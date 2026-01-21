VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frm_bio_metric_upload_new 
   Caption         =   "ATTENDANCE UPLOAD FROM BIO-METRIC SYSTEM"
   ClientHeight    =   8910
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14910
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8910
   ScaleWidth      =   14910
   WindowState     =   2  'Maximized
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   735
      Left            =   6720
      TabIndex        =   20
      Top             =   5880
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1296
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Frame frame_day 
      Height          =   975
      Left            =   1560
      TabIndex        =   17
      Top             =   3120
      Visible         =   0   'False
      Width           =   8295
      Begin MSComCtl2.DTPicker dt_ason 
         Height          =   375
         Left            =   3480
         TabIndex        =   19
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   59244545
         CurrentDate     =   39359
      End
      Begin VB.Label Label3 
         Caption         =   "Details for"
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
         Left            =   1920
         TabIndex        =   18
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1575
      Left            =   1560
      TabIndex        =   14
      Top             =   1080
      Width           =   8295
      Begin VB.OptionButton Opt2 
         Caption         =   "Attendance details upload for the month"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   480
         TabIndex        =   16
         Top             =   840
         Width           =   6975
      End
      Begin VB.OptionButton opt1 
         Caption         =   "Attendance details upload for the Day"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   480
         TabIndex        =   15
         Top             =   360
         Width           =   6975
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   4920
      TabIndex        =   10
      Top             =   4560
      Width           =   2175
      Begin VB.CommandButton exit 
         BackColor       =   &H00C0E0FF&
         Caption         =   "&Exit"
         Height          =   705
         Left            =   1080
         MaskColor       =   &H000000FF&
         Picture         =   "frm_bio_metric_upload_new.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton save 
         BackColor       =   &H00C0E0FF&
         Caption         =   "&Upload"
         Height          =   705
         Left            =   360
         MaskColor       =   &H000000FF&
         Picture         =   "frm_bio_metric_upload_new.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   120
         Width           =   735
      End
   End
   Begin VB.Frame frame_month 
      Height          =   1095
      Left            =   1560
      TabIndex        =   5
      Top             =   3360
      Visible         =   0   'False
      Width           =   8295
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
         TabIndex        =   7
         Top             =   480
         Width           =   3615
      End
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
         TabIndex        =   6
         Top             =   360
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
         TabIndex        =   9
         Top             =   480
         Width           =   885
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
         TabIndex        =   8
         Top             =   480
         Width           =   1335
      End
   End
   Begin VB.Frame Frame4 
      Height          =   975
      Left            =   720
      TabIndex        =   0
      Top             =   6000
      Visible         =   0   'False
      Width           =   3735
      Begin MSComCtl2.DTPicker st_date 
         Height          =   375
         Left            =   360
         TabIndex        =   1
         Top             =   480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   59244545
         CurrentDate     =   39359
      End
      Begin MSComCtl2.DTPicker end_date 
         Height          =   375
         Left            =   1920
         TabIndex        =   2
         Top             =   480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   59244545
         CurrentDate     =   39359
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
         TabIndex        =   4
         Top             =   240
         Width           =   1095
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
         TabIndex        =   3
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   5160
      Top             =   7200
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
      Left            =   240
      TabIndex        =   13
      Top             =   600
      Width           =   10695
   End
End
Attribute VB_Name = "frm_bio_metric_upload_new"
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
    dt_ason.Value = Now
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
''    With cmb_year
''''        .AddItem finyear + 2000
''        .AddItem "2012"
''        .AddItem "2013"
''        .AddItem "2014"
''        .AddItem "2015"
''        .Text = "2015"
''    End With
''
    With cmb_year
        .AddItem Left(fyear, 4)
        .AddItem Mid(fyear, 6, 4)
    End With

    ProgressBar1.Visible = False
    
End Sub

Private Sub opt1_Click()
    frame_day.Visible = True
    frame_month.Visible = False
End Sub

Private Sub Opt2_Click()
    frame_day.Visible = False
    frame_month.Visible = True
End Sub

Private Sub SAVE_Click()
On Error GoTo err_handler
    
    Dim id, fcode As Integer
    Dim dlogdate As Date
    
    Dim dev_log(100) As Long
    
    Dim log_details As String
    
    
    Dim stime, etime As Date
    
    stime = TimeValue(Now)
    
    
    Dim sft, sft_bt, sft_et, sft_begin_dur, sft_end_dur As String
    
    ProgressBar1.Visible = True
    
    If Opt2.Value = True Then
        If cmb_month.Text = "" Then
           MsgBox ("Select Month...")
           Exit Sub
        End If
        If cmb_year.Text = "" Then
           MsgBox ("Select Year...")
           Exit Sub
        End If
        find_dates
    Else
        end_date = dt_ason.Value
        st_date = dt_ason.Value
    
    End If
    
    ProgressBar1.Value = ProgressBar1.Min
    
    Dim dsnmdb As String
    Dim mdbrs As New ADODB.Recordset
    
    Set paydb = New ADODB.Connection
    Set payrs = New ADODB.Recordset
    
    paydb.Open pay

    dsnmdb = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=\\192.168.11.31\eSSL\eTimeTrackLite\eTimeTrackLite1.mdb"
    
    dsnmdb = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=c:\eTimeTrackLite1.mdb"


''    paydb.BeginTrans
'start

    pst_qry = "delete from bio_devicelogs where ad_logdate between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "'"
    pst_qry = "delete from bio_devicelogs"

    paydb.Execute pst_qry


    pst_qry = "delete from bio_device_shiftlogs where ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "'"
    pst_qry = "delete from bio_device_shiftlogs"
    paydb.Execute pst_qry


'''---select MSACESS MDB FILE

''    mdb_qry = "Select * from devicelogs as a, employees as b where a.deviceid <> 1 and a.userid = b.EmployeeCode and userid = '8007' and logdate between #" & Format(st_date, "MM/dd/yyyy") & "# and  #" & Format(end_date, "MM/dd/yyyy") & "# order by devicelogid"

    mdb_qry = "Select * from devicelogs as a, employees as b where a.deviceid <> 1 and a.userid = b.EmployeeCode  and logdate between #" & Format(st_date, "MM/dd/yyyy") & "# and  #" & Format(end_date, "MM/dd/yyyy") & "# order by devicelogid "

    mdbrs.Open mdb_qry, dsnmdb, adOpenDynamic, adLockOptimistic
    While Not mdbrs.EOF
'         ProgressBar1.Value = ProgressBar1.Value + 1
         pst_qry = "insert into  bio_devicelogs (ad_fpcode,ad_empid,ad_logslno,ad_date,ad_logdate) values (  " & mdbrs!userid & ", " & mdbrs!employeeid & "," & mdbrs!devicelogid & ", '" & Format(mdbrs!logdate, "MM/dd/yyyy") & "' , '" & Format(mdbrs!logdate, "MM/dd/yyyy HH:MM:SS") & "'  )"
         paydb.Execute pst_qry
         mdbrs.MoveNext
    Wend
    mdbrs.Close


    Dim idate, sft_from_date, sft_end_date As Date

    mdb_qry = "Select * from employees where employeecode <> '0' and employeeid = 2819 and Status = 'Working' "
    mdb_qry = "Select * from employees where employeecode <> '0' and Status = 'Working' "
    mdbrs.Open mdb_qry, dsnmdb, adOpenDynamic, adLockOptimistic
    While Not mdbrs.EOF
         For idate = st_date To end_date
            pst_qry = "insert into  bio_device_shiftlogs (ds_fpcode,ds_empid,ds_date) values (  " & mdbrs!employeecode & ", " & mdbrs!employeeid & ",  '" & Format(idate, "MM/dd/yyyy") & "'  )"
            paydb.Execute pst_qry
         Next
         mdbrs.MoveNext
    Wend
    mdbrs.Close

    Dim sftfound As Integer
    sql = "update bio_device_shiftlogs set ds_shift = 'GS',ds_shift_begintime = '08:00',ds_shift_endtime = '17:00',ds_begin_duration  = '60',ds_end_duration = '420'"
    paydb.Execute sql

''Weekoff updation for sundays

''    sql = "update bio_device_shiftlogs set ds_shift = 'WO',ds_shift_begintime = '',ds_shift_endtime = '',ds_begin_duration  = '',ds_end_duration = '' where UCase(Format(ds_date, ""dddd"")) = 'SUNDAY'"
''    paydb.Execute sql
    
    mdb_qry = "Select * from employeeshift as a, shifts as b where a.shiftid = b.shiftid  and fromdate >=  #" & Format(st_date, "dd/MM/yyyy") & "# and todate <=  #" & Format(end_date, "dd/MM/yyyy") & "# "
    mdbrs.Open mdb_qry, dsnmdb, adOpenDynamic, adLockOptimistic
    While Not mdbrs.EOF
          id = mdbrs!employeeid
          sft = mdbrs!shiftsname
          sft_bt = mdbrs!begintime
          sft_et = mdbrs!endtime
          sft_begin_dur = mdbrs!punchbeginduration
          sft_end_dur = mdbrs!punchendduration
          sft_from_date = mdbrs!fromdate
          sft_end_date = mdbrs!todate
          sql = "update bio_device_shiftlogs set ds_shift = '" & sft & "',ds_shift_begintime = '" & sft_bt & "',ds_shift_endtime = '" & sft_et & "',ds_begin_duration  = '" & sft_begin_dur & "',ds_end_duration = '" & sft_begin_dur & "' where ds_empid =  '" & id & "' and ds_date between '" & Format(sft_from_date, "MM/dd/yyyy") & "' and '" & Format(sft_end_date, "MM/dd/yyyy") & "' "
          paydb.Execute sql
          mdbrs.MoveNext
    Wend
    mdbrs.Close


''for updating general shift , A Shift , B Shift
    pst_qry = "Update bio_device_shiftlogs set ds_shift_in = intime, ds_shift_out = outtime  from bio_device_shiftlogs a , (select ad_fpcode,ad_empid,ad_date, min(ad_logdate) as intime , max(ad_logdate) as outtime    from bio_devicelogs  group by ad_fpcode,ad_empid,ad_date ) b Where ds_shift in ('GS','A SHIFT','B SHIFT','6.00 to 6.00 (Day)') and ds_fpcode = ad_fpcode And ds_empid = ad_empid And ds_date = ad_date"
    paydb.Execute pst_qry

    Set paydb2 = New ADODB.Connection
    Set payrs2 = New ADODB.Recordset
    paydb2.Open pay


    Dim i, j As Integer

    Dim dt_log As Integer


''for updating B Shift - 02.00 PM - 10.00 PM
    sql = "select * from bio_device_shiftlogs where ds_shift = 'B SHIFT'"
    payrs.Open sql, paydb, 1, 2
    While Not payrs.EOF
         dt_log = 1
          i = 1
          id = payrs!ds_empid
          sft_from_date = payrs!ds_date
          pst_qry = "select ad_date,ad_logslno,min(ad_logdate) as gtime from bio_devicelogs where ad_upd  = 'N' and ad_empid = " & id & " and ad_date >= '" & Format(sft_from_date, "MM/dd/yyyy") & "' and ad_date <= '" & Format(sft_from_date + 1, "MM/dd/yyyy") & "' and ad_upd = 'N' group by ad_date,ad_logslno"
          payrs2.Open pst_qry, paydb2, 1, 2
          While Not payrs2.EOF
              If i = 1 Then
                 payrs("ds_shift_in") = payrs2!gtime
                 dev_log(dt_log) = payrs2!ad_logslno
                 dt_log = dt_log + 1

              ElseIf i = 2 Then
                 payrs("ds_shift_out") = payrs2!gtime
                 dev_log(dt_log) = payrs2!ad_logslno
                 dt_log = dt_log + 1

              End If
              i = i + 1
              payrs2.MoveNext
          Wend
          payrs2.Close
          payrs.Update

          log_details = "(0"
          For j = 1 To dt_log - 1
              log_details = log_details + "," + Str(dev_log(j))
          Next
          log_details = log_details + ")"
          pst_qry = "update bio_devicelogs set ad_upd  = 'Y' where  ad_empid = " & id & " and ad_logslno in " & log_details
          paydb2.Execute pst_qry

          payrs.MoveNext


    Wend
    payrs.Close






''for updating B&C Shift - 6.00 PM - 6.00 AM and   C SHIFT - 2.00 PM - 6.00 AM
    sql = "select * from bio_device_shiftlogs where ds_shift = '6.00 to 6.00 (Night)'"

    sql = "select * from bio_device_shiftlogs where ds_shift in ('6.00 to 6.00 (Night)','C SHIFT')"
    payrs.Open sql, paydb, 1, 2
    While Not payrs.EOF
          dt_log = 1
          i = 1
          id = payrs!ds_empid
          sft_from_date = payrs!ds_date

''          pst_qry = "select ad_date,ad_logslno,min(ad_logdate) as gtime from bio_devicelogs where ad_empid = " & id & " and ad_date >= '" & Format(sft_from_date, "MM/dd/yyyy") & "' and ad_date <= '" & Format(sft_from_date + 1, "MM/dd/yyyy") & "' group by ad_date,ad_logslno"

          pst_qry = "select ad_date,ad_logslno,gtime from " _
              & " (select ad_date,ad_logslno,min(ad_logdate) as gtime from bio_devicelogs where ad_empid =  " & id & "  and ad_date = '" & Format(sft_from_date, "MM/dd/yyyy") & "' and datepart(hh,ad_logdate) > 15  and ad_upd = 'N'  group by ad_date,ad_logslno " _
              & " Union All " _
              & " select ad_date,ad_logslno,min(ad_logdate) as gtime from bio_devicelogs where ad_empid =  " & id & "  and ad_date = '" & Format(sft_from_date + 1, "MM/dd/yyyy") & "' and datepart(hh,ad_logdate) <7  and ad_upd = 'N' group by ad_date,ad_logslno ) a group by  ad_date,ad_logslno,gtime "

''          pst_qry = "select ad_date,ad_logslno,gtime,ad_upd from " _
''              & " (select ad_date,ad_logslno,min(ad_logdate) as gtime,ad_upd from bio_devicelogs where ad_empid =  " & id & "  and ad_date = '" & Format(sft_from_date, "MM/dd/yyyy") & "' and datepart(hh,ad_logdate) > 15  group by ad_date,ad_logslno,ad_upd " _
''              & " Union All " _
''              & " select ad_date,ad_logslno,min(ad_logdate) as gtime,ad_upd from bio_devicelogs where ad_empid =  " & id & "  and ad_date = '" & Format(sft_from_date + 1, "MM/dd/yyyy") & "' and datepart(hh,ad_logdate) <7  group by ad_date,ad_logslno,ad_upd ) a group by  ad_date,ad_logslno,gtime,ad_upd "


          payrs2.Open pst_qry, paydb2, 1, 2
          While Not payrs2.EOF
              If i = 1 Then
                 payrs("ds_shift_in") = payrs2!gtime
                 dev_log(dt_log) = payrs2!ad_logslno
                 dt_log = dt_log + 1

              ElseIf i = 2 Then
                 payrs("ds_shift_out") = payrs2!gtime
                 dev_log(dt_log) = payrs2!ad_logslno
                 dt_log = dt_log + 1
              End If
              i = i + 1

              payrs2.MoveNext
          Wend
          payrs2.Close
          payrs.Update

          log_details = "(0"
          For j = 1 To dt_log - 1
              log_details = log_details + "," + Str(dev_log(j))
          Next
          log_details = log_details + ")"
          pst_qry = "update bio_devicelogs set ad_upd  = 'Y' where  ad_empid = " & id & " and ad_logslno in " & log_details
          paydb2.Execute pst_qry
          payrs.MoveNext
    Wend
    payrs.Close

''for updating A+B+C Shift - 6.00 AM - NEXT DAY 6.00 AM
''    dt_log = 1

    sql = "select * from bio_device_shiftlogs where ds_shift = 'A+B+C'"
    payrs.Open sql, paydb, 1, 2
    While Not payrs.EOF
          dt_log = 1
          i = 1
          id = payrs!ds_empid
          sft_from_date = payrs!ds_date

''          pst_qry = "select ad_date,ad_logslno,min(ad_logdate) as gtime from bio_devicelogs where ad_empid = " & id & " and ad_date >= '" & Format(sft_from_date, "MM/dd/yyyy") & "' and ad_date <= '" & Format(sft_from_date + 1, "MM/dd/yyyy") & "' group by ad_date,ad_logslno"

          pst_qry = "select ad_date,ad_logslno,gtime from " _
              & " (select ad_date,ad_logslno,min(ad_logdate) as gtime from bio_devicelogs where ad_empid =  " & id & "  and ad_date = '" & Format(sft_from_date, "MM/dd/yyyy") & "' and datepart(hh,ad_logdate) > 5  and ad_upd = 'N' group by ad_date,ad_logslno " _
              & " Union All " _
              & " select ad_date,ad_logslno,min(ad_logdate) as gtime from bio_devicelogs where ad_empid =  " & id & "  and ad_date = '" & Format(sft_from_date + 1, "MM/dd/yyyy") & "' and datepart(hh,ad_logdate) <7  and ad_upd = 'N' group by ad_date,ad_logslno ) a group by  ad_date,ad_logslno,gtime "


          payrs2.Open pst_qry, paydb2, 1, 2
          While Not payrs2.EOF
              If i = 1 Then
                 payrs("ds_shift_in") = payrs2!gtime
                 dev_log(dt_log) = payrs2!ad_logslno
                 dt_log = dt_log + 1
              ElseIf i = 2 Then
                 payrs("ds_shift_out") = payrs2!gtime
                 dev_log(dt_log) = payrs2!ad_logslno
                 dt_log = dt_log + 1
              End If
              i = i + 1
  ''            payrs2("ad_upd") = "Y"
  ''            payrs2.Update
              payrs2.MoveNext
          Wend
          payrs2.Close
          payrs.Update
          log_details = "(0"
          For j = 1 To dt_log - 1
             log_details = log_details + "," + Str(dev_log(j))
          Next
          log_details = log_details + ")"
          pst_qry = "update bio_devicelogs set ad_upd  = 'Y' where  ad_empid = " & id & " and ad_logslno in " & log_details
          paydb2.Execute pst_qry

          payrs.MoveNext
    Wend
    payrs.Close
'end
    Dim leave, leavetype As String
 ''for updating leave entries
    mdb_qry = "Select * from leaveentries as a, leavetypes as b where a.leavetypeid = b.leavetypeid  and fromdate >=  #" & Format(st_date, "MM/dd/yyyy") & "# and todate <=  #" & Format(end_date, "MM/dd/yyyy") & "# order by employeeid , fromdate "
    mdbrs.Open mdb_qry, dsnmdb, adOpenDynamic, adLockOptimistic
    While Not mdbrs.EOF
           leave = ""
           If mdbrs!leavestatus = "HalfDay" Then
              leave = "½"
           End If
           sft_from_date = mdbrs!fromdate
           sft_end_date = mdbrs!todate
           leavetype = leave + mdbrs!leavetypesname
           id = mdbrs!employeeid
''           If id = 2642 Then
''              MsgBox ("Wait")
''           End If

           sql = "update bio_device_shiftlogs set ds_status = '" & leavetype & "' where ds_empid =  '" & id & "' and ds_date between '" & Format(sft_from_date, "MM/dd/yyyy") & "' and '" & Format(sft_end_date, "MM/dd/yyyy") & "' "
           paydb.Execute sql
           mdbrs.MoveNext
    Wend
    mdbrs.Close


 ''for updating OD entries
    mdb_qry = "Select * from specialentries where  fromdate >=  #" & Format(st_date, "MM/dd/yyyy") & "# and todate <=  #" & Format(end_date, "MM/dd/yyyy") & "# order by employeeid , fromdate "
    mdbrs.Open mdb_qry, dsnmdb, adOpenDynamic, adLockOptimistic
    While Not mdbrs.EOF
           sft_from_date = mdbrs!fromdate
           sft_end_date = mdbrs!todate
           id = mdbrs!employeeid
           sql = "update bio_device_shiftlogs set ds_status = 'P(OD)' where ds_empid =  '" & id & "' and ds_date between '" & Format(sft_from_date, "MM/dd/yyyy") & "' and '" & Format(sft_end_date, "MM/dd/yyyy") & "' "
           paydb.Execute sql
           mdbrs.MoveNext
    Wend
    mdbrs.Close

    
    
    
    
    sql = "update bio_device_shiftlogs set ds_sft_hrs = (case when DATEpart(minute,ds_shift_out) > 40 then datepart(hour,ds_shift_out)+1+datediff(day,ds_shift_in,ds_shift_out)*24  else datepart(hour,ds_shift_out)+datediff(day,ds_shift_in,ds_shift_out)*24 end - case when DATEpart(minute,ds_shift_in) > 30 then datepart(hour,ds_shift_in)+1 else datepart(hour,ds_shift_in) end)"
    paydb.Execute sql
 
    sql = "update bio_device_shiftlogs set ds_status = 'P' WHERE ds_sft_hrs > 7"
    paydb.Execute sql
 
    sql = "update bio_device_shiftlogs set ds_status = 'WO' WHERE   ds_shift = 'WO'"
    paydb.Execute sql
 
 
    etime = TimeValue(Now)
 
    MsgBox ("Process Start by " + Str(stime) + " ...  end by " + Str(etime))
    
 
 
 

    MsgBox ("Updated...")
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
    If ProgressBar1.Value + 10 > ProgressBar1.Max Then
         ProgressBar1.Max = ProgressBar1.Max + 100
    End If
    ProgressBar1.Value = (ProgressBar1.Value + 1)
End Sub

