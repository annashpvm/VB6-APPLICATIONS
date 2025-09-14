VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form frm_bio_metric_upload 
   Caption         =   "ATTENDANCE UPLOAD FROM BIO-METRIC"
   ClientHeight    =   8955
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12345
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8955
   ScaleWidth      =   12345
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   5760
      Top             =   4920
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   495
      Left            =   1560
      TabIndex        =   14
      Top             =   4200
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   873
      _Version        =   393216
      Appearance      =   1
      Max             =   5000
   End
   Begin VB.Frame Frame4 
      Height          =   975
      Left            =   840
      TabIndex        =   9
      Top             =   6000
      Visible         =   0   'False
      Width           =   3735
      Begin MSComCtl2.DTPicker st_date 
         Height          =   375
         Left            =   360
         TabIndex        =   10
         Top             =   480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   16515073
         CurrentDate     =   39359
      End
      Begin MSComCtl2.DTPicker end_date 
         Height          =   375
         Left            =   1920
         TabIndex        =   11
         Top             =   480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   16515073
         CurrentDate     =   39359
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
         TabIndex        =   13
         Top             =   240
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
         Left            =   2040
         TabIndex        =   12
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1935
      Left            =   1680
      TabIndex        =   4
      Top             =   1200
      Width           =   8055
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
         Left            =   5280
         TabIndex        =   6
         Top             =   840
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
         Left            =   840
         TabIndex        =   5
         Top             =   840
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
         Left            =   840
         TabIndex        =   8
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
         Left            =   5520
         TabIndex        =   7
         Top             =   480
         Width           =   885
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   5040
      TabIndex        =   1
      Top             =   3240
      Width           =   2175
      Begin VB.CommandButton save 
         BackColor       =   &H00C0E0FF&
         Caption         =   "&Upload"
         Height          =   705
         Left            =   360
         MaskColor       =   &H000000FF&
         Picture         =   "frm_bio_metric_upload.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton exit 
         BackColor       =   &H00C0E0FF&
         Caption         =   "&Exit"
         Height          =   705
         Left            =   1080
         MaskColor       =   &H000000FF&
         Picture         =   "frm_bio_metric_upload.frx":066A
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   120
         Width           =   735
      End
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
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   10695
   End
End
Attribute VB_Name = "frm_bio_metric_upload"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmb_month_Change()
   find_dates
End Sub

Private Sub cmb_year_Change()
    find_dates
End Sub

Private Sub exit_Click()
    Unload Me
End Sub

Private Sub Form_Load()
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
''        .AddItem finyear + 2000
        .AddItem "2012"
        .AddItem "2013"
        .AddItem "2014"
        .AddItem "2015"
        .AddItem "2016"

        .Text = "2015"
    End With
    ProgressBar1.Visible = False
    
End Sub

Private Sub save_Click()
On Error GoTo err_handler
    
    ProgressBar1.Visible = True
    If cmb_month.Text = "" Then
       MsgBox ("Select Month...")
       Exit Sub
    End If
    If cmb_year.Text = "" Then
       MsgBox ("Select Year...")
       Exit Sub
    End If
    find_dates
    ProgressBar1.Value = ProgressBar1.Min
    Dim dsnmdb As String
    Dim mdbrs As New ADODB.Recordset

''    paydb.BeginTrans
    
    pst_qry = "delete from bio_attendlogs where a_year = " & Val(cmb_year.Text) & "  and a_month = " & cmb_month.ItemData(cmb_month.ListIndex)
    paydb.Execute pst_qry
    
    pst_qry = "delete from bio_attendlogs_daily where a_year = " & Val(cmb_year.Text) & "  and a_month = " & cmb_month.ItemData(cmb_month.ListIndex)
    paydb.Execute pst_qry


'''---select MSACESS MDB FILE
''    dsnmdb = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=d:\eTimeTrackLite1.mdb"
                                                                
    dsnmdb = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=\\192.168.11.31\eSSL\eTimeTrackLite\eTimeTrackLite1.mdb"
''    dsnmdb = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=c:\eTimeTrackLite1.mdb"

''        mdb_qry = "Select a.EmployeeId,b.employeecode from attendancelogs as a, employees as b where a.EmployeeId =  b.EmployeeId  and a.Attendancedate between #" & Format(st_date, "MM/dd/yyyy") & "# and  #" & Format(end_date, "MM/dd/yyyy") & "# and b.Status = 'Working'  and b.employeecode = '5012'  group by a.EmployeeId,b.employeecode"
    mdb_qry = "Select a.EmployeeId,b.employeecode from attendancelogs as a, employees as b where a.EmployeeId =  b.EmployeeId  and a.Attendancedate between #" & Format(st_date, "MM/dd/yyyy") & "# and  #" & Format(end_date, "MM/dd/yyyy") & "# and b.Status = 'Working'  group by a.EmployeeId,b.employeecode"
    mdbrs.Open mdb_qry, dsnmdb, adOpenDynamic, adLockOptimistic
''    ProgressBar1.Max = mdbrs.RecordCount
    While Not mdbrs.EOF
         ProgressBar1.Value = ProgressBar1.Value + 1
         pst_qry = "insert into bio_attendlogs (a_bioid,a_fpcode,a_month,a_year) values (  " & mdbrs!employeeid & "," & mdbrs!employeecode & ", " & cmb_month.ItemData(cmb_month.ListIndex) & " , " & Val(cmb_year.Text) & " )"
         paydb.Execute pst_qry
         mdbrs.MoveNext
    Wend
    mdbrs.Close





    mdb_qry = "Select a.EmployeeId,b.employeecode,a.attendancedate,a.statuscode,a.intime,a.outtime,a.lateby,a.earlyby,a.overtime from attendancelogs as a, employees as b where a.EmployeeId =  b.EmployeeId  and a.Attendancedate between #" & Format(st_date, "MM/dd/yyyy") & "# and  #" & Format(end_date, "MM/dd/yyyy") & "# and b.Status = 'Working'  group by a.EmployeeId,b.employeecode,a.attendancedate,a.statuscode,a.intime,a.outtime,a.lateby,a.earlyby,a.overtime"
    mdbrs.Open mdb_qry, dsnmdb, adOpenDynamic, adLockOptimistic
    While Not mdbrs.EOF
         ''ProgressBar1.Value = ProgressBar1.Value + 1
         pst_qry = "insert into bio_attendlogs_daily (a_bioid,a_fpcode,a_month,a_year,a_status,a_date,a_date_in,a_date_out,a_lateby,a_earlyby,a_overtime) values (  " & mdbrs!employeeid & "," & mdbrs!employeecode & ", " & cmb_month.ItemData(cmb_month.ListIndex) & " , " & Val(cmb_year.Text) & ",'" & mdbrs!StatusCode & "','" & Format(mdbrs!Attendancedate, "MM/dd/yyyy") & "','" & mdbrs!intime & "','" & mdbrs!outtime & "','" & mdbrs!lateby & "','" & mdbrs!earlyby & "','" & mdbrs!overtime & "' )"
         paydb.Execute pst_qry
         mdbrs.MoveNext
    Wend
    mdbrs.Close
    
    
    
    
    
    Dim aday As String
''        mdb_qry = "Select * from attendancelogs as a, employees as b where a.EmployeeId =  b.EmployeeId  and b.employeecode = '1042' and a.Attendancedate between #" & Format(st_date, "MM/dd/yyyy") & "# and  #" & Format(end_date, "MM/dd/yyyy") & "#  order by a.Attendancedate"
''        mdb_qry = "Select * from attendancelogs where EmployeeId = 2642 and   Attendancedate between #" & Format(st_date, "MM/dd/yyyy") & "# and  #" & Format(end_date, "MM/dd/yyyy") & "#  order by Attendancedate"
''        mdb_qry = "Select * from attendancelogs where EmployeeId = 2643 and  Attendancedate between #" & Format(st_date, "MM/dd/yyyy") & "# and  #" & Format(end_date, "MM/dd/yyyy") & "#   order by Attendancedate"
    mdb_qry = "Select * from attendancelogs where Attendancedate between #" & Format(st_date, "MM/dd/yyyy") & "# and  #" & Format(end_date, "MM/dd/yyyy") & "#   order by Attendancedate"
    mdbrs.Open mdb_qry, dsnmdb, adOpenDynamic, adLockOptimistic
    While Not mdbrs.EOF
  ''       ProgressBar1.Value = ProgressBar1.Value + 1
         aday = Trim(Str(Day(mdbrs!Attendancedate)))
         pst_qry = "update bio_attendlogs set a_day" & aday & " = '" & mdbrs!StatusCode & "',a_in_day" & aday & " = '" & mdbrs!intime & "' ,a_out_day" & aday & " = '" & mdbrs!outtime & "' where a_bioid = " & mdbrs!employeeid & " and a_year = " & Val(cmb_year.Text) & "  and a_month = " & cmb_month.ItemData(cmb_month.ListIndex)
         paydb.Execute pst_qry
         mdbrs.MoveNext
    Wend
    mdbrs.Close
    Dim dayfind, dayfind_intime, dayfind_outtime As String
    Dim present, absent, hop, wop, cl, sl, h, ch, layoff, wo, pl As Single
    Dim intime, outtime, difftime As Integer
    Set paydb = New ADODB.Connection
    Set payrs = New ADODB.Recordset
    sql = "select * from bio_attendlogs where a_year = " & Val(cmb_year.Text) & "  and a_month = " & cmb_month.ItemData(cmb_month.ListIndex)
    paydb.Open pay
    payrs.Open sql, paydb, 1, 2
    If Not payrs.EOF Then
       While Not payrs.EOF
             For i = 1 To 31
                dayfind = "a_day" & i
                dayfind_intime = "a_in_day" & i
                dayfind_outtime = "a_out_day" & i
                If IsNull(payrs.Fields(dayfind_intime)) = False And IsNull(payrs.Fields(dayfind_outtime)) = False Then
                    intime = Hour(payrs.Fields(dayfind_intime)) * 60 + Minute(payrs.Fields(dayfind_intime))
                    outtime = Hour(payrs.Fields(dayfind_outtime)) * 60 + Minute(payrs.Fields(dayfind_outtime))
                    difftime = outtime - intime
                    If payrs.Fields(dayfind) = "P" And difftime > 180 And difftime < 420 Then
                        payrs.Fields(dayfind) = "½P"
                        payrs.Update
                    End If
                End If
             Next
             payrs.MoveNext
        Wend
     Else
        MsgBox ("Details not available for the date ")
     End If
     payrs.Close
''        sql = "select * from bio_attendlogs where a_fpcode = 5012"
    sql = "select * from bio_attendlogs where a_year = " & Val(cmb_year.Text) & "  and a_month = " & cmb_month.ItemData(cmb_month.ListIndex)
    payrs.Open sql, paydb, 1, 2
    If Not payrs.EOF Then
       While Not payrs.EOF
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
       
            For i = 1 To 31
                dayfind = "a_day" & i
                If payrs.Fields(dayfind) = "P" Or payrs.Fields(dayfind) = "P(OD)" Or payrs.Fields(dayfind) = "½P(OD)" Or payrs.Fields(dayfind) = "A(OD)" Then
                    present = present + 1
                ElseIf payrs.Fields(dayfind) = "A" Then
                    absent = absent + 1
                ElseIf payrs.Fields(dayfind) = "PL" Or payrs.Fields(dayfind) = "PLP" Then
                    pl = pl + 1
                ElseIf payrs.Fields(dayfind) = "½PL" Then
                    pl = pl + 0.5
                    absent = absent + 0.5
                ElseIf payrs.Fields(dayfind) = "½PLP" Then
                    pl = pl + 0.5
                    present = present + 0.5
                ElseIf payrs.Fields(dayfind) = "CL" Or payrs.Fields(dayfind) = "CL½P" Or payrs.Fields(dayfind) = "CLP" Then
                    cl = cl + 1
                ElseIf payrs.Fields(dayfind) = "½CL" Then
                    absent = absent + 0.5
                    cl = cl + 0.5
                ElseIf payrs.Fields(dayfind) = "½CLP" Or payrs.Fields(dayfind) = "½CL½P" Then
                    present = present + 0.5
                    cl = cl + 0.5
                ElseIf payrs.Fields(dayfind) = "SL" Or payrs.Fields(dayfind) = "SLP" Then
                    sl = sl + 1
                ElseIf payrs.Fields(dayfind) = "½SLP" Then
                    sl = sl + 0.5
                    present = present + 0.5
                ElseIf payrs.Fields(dayfind) = "H" Then
                    h = h + 1
                ElseIf payrs.Fields(dayfind) = "½P" Then
                    present = present + 0.5
                    absent = absent + 0.5
                ElseIf payrs.Fields(dayfind) = "Layoff" Or payrs.Fields(dayfind) = "LayoffP" Then
                    layoff = layoff + 1
                ElseIf payrs.Fields(dayfind) = "C.H" Or payrs.Fields(dayfind) = "C.H½P" Or payrs.Fields(dayfind) = "C.HP" Or payrs.Fields(dayfind) = "C.HP(OD)" Then
                    ch = ch + 1
                ElseIf payrs.Fields(dayfind) = "HOP" Or payrs.Fields(dayfind) = "H½P(OD)" Then
                    hop = hop + 1
                ElseIf payrs.Fields(dayfind) = "WOP" Or payrs.Fields(dayfind) = "WOP(OD)" Or payrs.Fields(dayfind) = "WO(OD)" Then
                    wop = wop + 1
                ElseIf payrs.Fields(dayfind) = "WO" Or payrs.Fields(dayfind) = "WO½P" Then
                    wo = wo + 1
                ElseIf payrs.Fields(dayfind) = "½C.H" Then
                    ch = ch + 0.5
                    absent = absent + 0.5
                End If
            Next
            
            payrs("a_present") = present
            payrs("a_hop") = hop
            payrs("a_wop") = wop
            payrs("a_el") = cl
            payrs("a_pl") = pl
            payrs("a_ml") = sl
            payrs("a_holiday") = h
            payrs("a_ch") = ch
            payrs("a_layoff") = layoff
            payrs("a_absent") = absent
            payrs("a_wo") = wo
            payrs.Update
            payrs.MoveNext
        Wend
     Else
        MsgBox ("Details not available for the date ")
     End If
     payrs.Close
  ''   paydb.CommitTrans
     MsgBox ("Updated...")
     ProgressBar1.Visible = False
     ProgressBar1.Value = 0
     Exit Sub
err_handler:
     chk = gen_Validation(Err.Number, Err.Description)
    '' paydb.RollbackTrans
     Me.MousePointer = 1
  '  chk = gen_Validation(Err.Number, Err.Description)
      If chk = 1 Then Resume
     
End Sub

Public Sub find_dates()
    If cmb_month.ListIndex = -1 Then Exit Sub
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
    end_date = DateValue(Str(mmon) + "/" + Str(mdays) + "/" + cmb_year.Text)
    st_date = end_date - Day(end_date) + 1
End Sub


Private Sub Timer1_Timer()
    ProgressBar1.Value = (ProgressBar1.Value + 1)
End Sub
