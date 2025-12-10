VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frm_bio_metric_upload_new 
   Caption         =   "ATTENDANCE UPLOAD FROM BIO-METRIC SYSTEM - forthe month"
   ClientHeight    =   8910
   ClientLeft      =   270
   ClientTop       =   240
   ClientWidth     =   14910
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8910
   ScaleWidth      =   14910
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmd_ok 
      Caption         =   "Accept"
      Height          =   195
      Left            =   240
      TabIndex        =   27
      Top             =   1680
      Width           =   975
   End
   Begin VB.TextBox txt_pw 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   240
      PasswordChar    =   "#"
      TabIndex        =   26
      Top             =   1200
      Width           =   1095
   End
   Begin VB.CommandButton cmd_saturday 
      Caption         =   "SATURDAY CUT"
      Height          =   375
      Left            =   11280
      TabIndex        =   25
      Top             =   3720
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Frame Frame3 
      Height          =   2175
      Left            =   10200
      TabIndex        =   22
      Top             =   1080
      Visible         =   0   'False
      Width           =   1695
      Begin VB.OptionButton Option1 
         Caption         =   "Old Method"
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
         Left            =   240
         TabIndex        =   24
         Top             =   1080
         Width           =   1095
      End
      Begin VB.OptionButton opt_new 
         Caption         =   "New Method"
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
         Left            =   240
         TabIndex        =   23
         Top             =   360
         Value           =   -1  'True
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      Enabled         =   0   'False
      Height          =   2415
      Left            =   1680
      TabIndex        =   18
      Top             =   960
      Width           =   8415
      Begin VB.CommandButton cmd_update_all 
         Caption         =   "UPDATE ALL"
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
         Left            =   1080
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   1680
         Visible         =   0   'False
         Width           =   6255
      End
      Begin VB.CommandButton cmd_update_leave_od 
         Caption         =   "UPDATE LEAVE /CH/ OD PARTICULARS && PROCESS"
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
         Left            =   1080
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   960
         Width           =   6255
      End
      Begin VB.CommandButton cmd_upload_biometric 
         Caption         =   "UPLOAD BIOMETRIC ATTENDANCE LOGS"
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
         Left            =   1080
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   240
         Width           =   6255
      End
   End
   Begin VB.CommandButton cmd_shiftupdation 
      Caption         =   "Shift schedule UPDATION"
      Height          =   495
      Left            =   11280
      TabIndex        =   17
      Top             =   5640
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton cmd_wo_upd 
      Caption         =   "WO UPDATION"
      Height          =   495
      Left            =   11280
      TabIndex        =   16
      Top             =   5040
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton cmd_leave_updation 
      Caption         =   "LEAVE && OD UPDATION"
      Height          =   495
      Left            =   11280
      TabIndex        =   15
      Top             =   4440
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Frame frame_month 
      Enabled         =   0   'False
      Height          =   1095
      Left            =   1800
      TabIndex        =   6
      Top             =   3480
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
         TabIndex        =   8
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
         TabIndex        =   7
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
         TabIndex        =   10
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
         TabIndex        =   9
         Top             =   480
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   4920
      TabIndex        =   11
      Top             =   4560
      Width           =   2175
      Begin VB.CommandButton cmd_exit 
         BackColor       =   &H00C0E0FF&
         Caption         =   "&Exit"
         Height          =   705
         Left            =   1080
         MaskColor       =   &H000000FF&
         Picture         =   "frm_bio_metric_upload_new.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton save 
         BackColor       =   &H00C0E0FF&
         Caption         =   "&Upload"
         Height          =   705
         Left            =   240
         MaskColor       =   &H000000FF&
         Picture         =   "frm_bio_metric_upload_new.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   120
         Visible         =   0   'False
         Width           =   735
      End
   End
   Begin VB.Frame Frame4 
      Enabled         =   0   'False
      Height          =   975
      Left            =   720
      TabIndex        =   1
      Top             =   6000
      Width           =   3735
      Begin MSComCtl2.DTPicker st_date 
         Height          =   375
         Left            =   360
         TabIndex        =   2
         Top             =   480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   60555265
         CurrentDate     =   39359
      End
      Begin MSComCtl2.DTPicker end_date 
         Height          =   375
         Left            =   1920
         TabIndex        =   3
         Top             =   480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   60555265
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
         TabIndex        =   5
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
         TabIndex        =   4
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   5160
      Top             =   7200
   End
   Begin VB.PictureBox ProgressBar1 
      Height          =   495
      Left            =   960
      ScaleHeight     =   435
      ScaleWidth      =   8115
      TabIndex        =   0
      Top             =   7440
      Visible         =   0   'False
      Width           =   8175
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
      TabIndex        =   14
      Top             =   240
      Width           =   10695
   End
End
Attribute VB_Name = "frm_bio_metric_upload_new"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim all_chk As Integer

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

Private Sub cmd_leave_updation_Click()
''        If cmb_month.Text = "" Then
''           MsgBox ("Select Month...")
''           Exit Sub
''        End If
''        If cmb_year.Text = "" Then
''           MsgBox ("Select Year...")
''           Exit Sub
''        End If
''        find_dates
''
''
''
''    Dim dsnmdb As String
''    Dim mdbrs As New ADODB.Recordset
''
''    Set paydb = New ADODB.Connection
''    Set payrs = New ADODB.Recordset
''
''    paydb.Open pay
''
''    dsnmdb = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=\\192.168.11.31\eSSL\eTimeTrackLite\eTimeTrackLite1.mdb"
''    dsnmdb = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=\\192.168.11.23\d\ESSL\etimetracklite\eTimeTrackLite1.mdb"
''
''''    dsnmdb = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=c:\eTimeTrackLite1.mdb"
''
'''''---select MSACESS MDB FILE
''
''''    mdb_qry = "Select * from devicelogs as a, employees as b where a.deviceid <> 1 and a.userid = b.EmployeeCode and userid = '8007' and logdate between #" & Format(st_date, "MM/dd/yyyy") & "# and  #" & Format(end_date, "MM/dd/yyyy") & "# order by devicelogid"
''
''
''
''    pst_qry = "delete from bio_empleave where emp_leave_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "'"
''    paydb.Execute pst_qry
''
''    pst_qry = "delete from bio_emp_oddetails where empod_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "'"
''    paydb.Execute pst_qry
''
''
''
'''' for updating leave entries
''    Dim l1 As String
''    Dim idate As Date
''    mdb_qry = "Select * from leaveentries as a, leavetypes as b ,employees as c  where  a.employeeid = c.employeeid and  a.leavetypeid = b.leavetypeid  and fromdate >=  #" & Format(st_date, "MM/dd/yyyy") & "# and todate <=  #" & Format(end_date, "MM/dd/yyyy") & "# order by a.employeeid , fromdate "
''    mdbrs.Open mdb_qry, dsnmdb, adOpenDynamic, adLockOptimistic
''    While Not mdbrs.EOF
''           leave = ""
''           l1 = "F"
''           If mdbrs!leavestatus = "HalfDay" Then
''              leave = "½"
''              l1 = "1"
''           End If
''           sft_from_date = mdbrs!fromdate
''           sft_end_date = mdbrs!todate
''           leavetype = leave + mdbrs!leavetypesname
''           id = mdbrs!employeecode
''           For idate = sft_from_date To sft_end_date
''               pst_qry = "insert into  bio_empleave  (emp_leave_no,emp_fpcode,emp_leave_type,emp_leave_date,emp_leave_period) values ( 0, " & mdbrs!employeecode & ", '" & leavetype & "',  '" & Format(idate, "MM/dd/yyyy") & "','" & l1 & "'  )"
''               paydb.Execute pst_qry
''
''           Next
''          mdbrs.MoveNext
''    Wend
''    mdbrs.Close
''
''    Dim btime, etime As String
'' ''for updating OD entries
''    mdb_qry = "Select * from specialentries as a ,employees as c  where  a.employeeid = c.employeeid and  fromdate >=  #" & Format(st_date, "MM/dd/yyyy") & "# and todate <=  #" & Format(end_date, "MM/dd/yyyy") & "# order by a.employeeid , fromdate "
''    mdbrs.Open mdb_qry, dsnmdb, adOpenDynamic, adLockOptimistic
''    While Not mdbrs.EOF
''           sft_from_date = mdbrs!fromdate
''           sft_end_date = mdbrs!todate
''           id = mdbrs!employeecode
''           For idate = sft_from_date To sft_end_date
''               pst_qry = "insert into  bio_emp_oddetails (empod_no,empod_fpcode,empod_date,empod_fromtime,empod_totime,empod_location,empod_purpose) values ( 0, " & mdbrs!employeecode & ",  '" & Format(idate, "MM/dd/yyyy") & "', '" & mdbrs!begintime & "', '" & mdbrs!endtime & "','',''  )"
''               paydb.Execute pst_qry
''           Next
''           mdbrs.MoveNext
''    Wend
''    mdbrs.Close
''
''    MsgBox ("updated..")
''
''    paydb.Close
End Sub


Private Sub cmd_ok_Click()
    If txt_pw.Text = "ok" Then
       Frame1.Enabled = True
       Frame2.Enabled = True
       frame_month.Enabled = True
    End If
End Sub

Private Sub cmd_saturday_Click()
    pst_qry = "update bio_device_shiftlogs set ds_status = 'C' from bio_device_shiftlogs ,bio_emp_saturday where ds_fpcode = empsat_fpcode and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and '" & Format(end_date, "MM/dd/yyyy") & "' and datepart(dw,ds_date) =  7 and ds_status <> 'H'"
    paydb.Execute pst_qry

End Sub

Private Sub cmd_shiftupdation_Click()
        If cmb_month.Text = "" Then
           MsgBox ("Select Month...")
           Exit Sub
        End If
        If cmb_year.Text = "" Then
           MsgBox ("Select Year...")
           Exit Sub
        End If
        find_dates


    
    Dim dsnmdb As String
    Dim mdbrs As New ADODB.Recordset
    
    Set paydb = New ADODB.Connection
    Set payrs = New ADODB.Recordset
    
    paydb.Open pay

    dsnmdb = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=\\192.168.11.31\eSSL\eTimeTrackLite\eTimeTrackLite1.mdb"
    dsnmdb = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=\\192.168.11.23\d\ESSL\etimetracklite\eTimeTrackLite1.mdb"

''    dsnmdb = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=c:\eTimeTrackLite1.mdb"

'''---select MSACESS MDB FILE

''    mdb_qry = "Select * from devicelogs as a, employees as b where a.deviceid <> 1 and a.userid = b.EmployeeCode and userid = '8007' and logdate between #" & Format(st_date, "MM/dd/yyyy") & "# and  #" & Format(end_date, "MM/dd/yyyy") & "# order by devicelogid"






'' for updating WO DATE
    
    Dim leave As String
    Dim empid As String
    
    Dim idate As Date
''Employee shift updation
    mdb_qry = "Select * from employeeshift as a, shifts as b ,employees c where c.Employeeid = a.Employeeid  and  a.shiftid = b.shiftid  and fromdate >=  #" & Format(st_date, "MM/dd/yyyy") & "# and todate <=  #" & Format(end_date, "MM/dd/yyyy") & "# "
    mdbrs.Open mdb_qry, dsnmdb, adOpenDynamic, adLockOptimistic
    While Not mdbrs.EOF
          id = mdbrs!employeecode
          sft = mdbrs!shiftsname
          sft_bt = mdbrs!begintime
          sft_et = mdbrs!endtime
          sft_begin_dur = mdbrs!punchbeginduration
          sft_end_dur = mdbrs!punchendduration
          sft_from_date = mdbrs!fromdate
          sft_end_date = mdbrs!todate
          For idate = sft_from_date To sft_end_date

''          sql = "update bio_device_shiftlogs set ds_shift = '" & sft & "',ds_shift_begintime = '" & sft_bt & "',ds_shift_endtime = '" & sft_et & "',ds_begin_duration  = '" & sft_begin_dur & "',ds_end_duration = '" & sft_begin_dur & "' where ds_empid =  '" & id & "' and ds_date between '" & Format(sft_from_date, "MM/dd/yyyy") & "' and '" & Format(sft_end_date, "MM/dd/yyyy") & "' "
              sql = "insert into  bio_shift_schedule  (emps_fpcode,emps_date,emps_shift_alloted,emps_shift) values ( " & mdbrs!employeecode & ", '" & Format(idate, "MM/dd/yyyy") & "','" & sft & "','" & sft & "')"
              paydb.Execute sql
          Next
          mdbrs.MoveNext
    Wend
    mdbrs.Close
    
    
    MsgBox ("updated..")

    paydb.Close


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

''
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



''      sql = "update bio_device_shiftlogs set  ds_shift =  emps_shift  , ds_shift_actual  =  emps_shift from bio_device_shiftlogs a, bio_shift_schedule b where emps_date = ds_date and  emps_fpcode = ds_fpcode and  emps_date between '" & Format(st_date, "MM/dd/yyyy") & "' and '" & Format(end_date, "MM/dd/yyyy") & "' "
''      paydb.Execute sql
  


        Set paydb2 = New ADODB.Connection
        Set payrs2 = New ADODB.Recordset
        paydb2.Open pay
        
        Dim i, j As Integer
        Dim dt_log As Integer
        Dim inpunch_chk, outpunch_chk, incount, outcount   As Integer


       Dim rs_set As New ADODB.Recordset
       Dim newqry As String
       
    ''Annadurai
    end_date = end_date - 1
''    For idate = st_date - 1 To end_date
        
''        If idate = "06/11/2020" Then
''           MsgBox ("TEst")
''        End If
        
''        sql = "select * from bio_devicelogs where ad_fpcode = 1018 and ad_date between '11/01/2020' and '11/30/2020' order by ad_date,ad_logdate"
        
        sql = "select * from bio_device_shiftlogs where ds_fpcode = 1018 and  ds_date  between  '" & Format(st_date.Value, "MM/dd/yyyy") & "' and '" & Format(end_date.Value, "MM/dd/yyyy") & "' and ds_fpcode > 0  order by ds_fpcode,ds_date"
''        sql = "select * from bio_device_shiftlogs where  ds_date  = '" & Format(idate, "MM/dd/yyyy") & "' and ds_fpcode > 0  order by ds_fpcode,ds_date"
        newqry = "select * from bio_device_shiftlogs where ds_date  between  '" & Format(st_date.Value, "MM/dd/yyyy") & "' and '" & Format(end_date.Value, "MM/dd/yyyy") & "' and ds_fpcode > 0  order by ds_fpcode,ds_date"
        
        newqry = "select ds_fpcode,ds_date from bio_device_shiftlogs,bio_devicelogs where ds_fpcode = ad_fpcode and ds_date = ad_date and ds_date  between  '" & Format(st_date.Value, "MM/dd/yyyy") & "' and '" & Format(end_date.Value, "MM/dd/yyyy") & "' and ds_fpcode > 0  group by ds_fpcode,ds_date order by ds_fpcode,ds_date"
        
''        newqry = "select ds_fpcode,ds_date from bio_device_shiftlogs,bio_devicelogs where ds_fpcode = 1018 and ds_fpcode = ad_fpcode  and ds_date  between  '" & Format(st_date.Value, "MM/dd/yyyy") & "' and '" & Format(end_date.Value, "MM/dd/yyyy") & "' and ds_fpcode > 0  group by ds_fpcode,ds_date order by ds_fpcode,ds_date"
        
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
''              If sft_from_date = "03/11/2020" Then
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
                             ElseIf incount = 2 Then
                                 payrs("ds_shift_in2") = payrs2!ad_logdate
                                 incount = incount + 1
                             ElseIf incount = 3 Then
                                 payrs("ds_shift_in3") = payrs2!ad_logdate
                                  incount = incount + 1
                             ElseIf incount = 4 Then
                                 payrs("ds_shift_in4") = payrs2!ad_logdate
                                 incount = incount + 1
                             ElseIf incount = 5 Then
                                 payrs("ds_shift_in5") = payrs2!ad_logdate
                                 incount = incount + 1
                             ElseIf incount = 6 Then
                                 payrs("ds_shift_in6") = payrs2!ad_logdate
                                 incount = incount + 1
                             End If
                             inpunch_chk = 1
                    ElseIf Trim(payrs2!ad_punch) = "out" Then
                             If outcount = 1 Then
                                 If payrs2!ad_logdate > payrs("ds_shift_in") And payrs("ds_shift_in") <> "01/01/1900" Then
                                    payrs("ds_shift_out") = payrs2!ad_logdate
                                    payrs("ds_no_of_punches") = outcount
                                    outcount = outcount + 1
                                    outpunch_chk = 1
                                 End If
                             ElseIf outcount = 2 Then
                                  If payrs2!ad_logdate > payrs("ds_shift_in2") And payrs("ds_shift_in2") <> "01/01/1900" Then
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
                    dev_log(dt_log) = payrs2!ad_logslno
                    dt_log = dt_log + 1
                    payrs2.MoveNext
              Wend
              payrs2.Close
''              If incount > outcount Then
''                    pst_qry = "select * from bio_devicelogs where ad_fpcode =  " & id & " and ad_punch = 'out'  and ad_date = '" & Format(sft_from_date + 1, "MM/dd/yyyy") & "'  and ad_upd = 'N' and datepart(hh,ad_logdate) <= 9   order by ad_date,ad_logdate"
''                    payrs2.Open pst_qry, paydb2, 1, 2
''                    While Not payrs2.EOF
''
''                          If payrs2!ad_punch = "out" Then
''                             If outcount = 1 Then
''                                 payrs("ds_shift_out") = payrs2!ad_logdate
''                                  payrs("ds_no_of_punches") = outcount
''                                 outcount = outcount + 1
''                             ElseIf outcount = 2 Then
''                                 payrs("ds_shift_out2") = payrs2!ad_logdate
''                                 payrs("ds_no_of_punches") = outcount
''                                 outcount = outcount + 1
''                             ElseIf outcount = 3 Then
''                                 payrs("ds_shift_out3") = payrs2!ad_logdate
''                                 payrs("ds_no_of_punches") = outcount
''                                 outcount = outcount + 1
''                             ElseIf outcount = 4 Then
''                                 payrs("ds_shift_out4") = payrs2!ad_logdate
''                                 payrs("ds_no_of_punches") = outcount
''                                 outcount = outcount + 1
''                             ElseIf outcount = 5 Then
''                                 payrs("ds_shift_out5") = payrs2!ad_logdate
''                                 payrs("ds_no_of_punches") = outcount
''                                 outcount = outcount + 1
''                             ElseIf incount = 6 Then
''                                 payrs("ds_shift_out6") = payrs2!ad_logdate
''                                 payrs("ds_no_of_punches") = outcount
''                                 outcount = outcount + 1
''                             End If
''
''                             dev_log(dt_log) = payrs2!ad_logslno
''                             dt_log = dt_log + 1
''                          End If
''                          payrs2.MoveNext
''                    Wend
''                    payrs2.Close
''
''              End If
              
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
        


        
        
150:
    paydb.CommandTimeout = 300
    sql = "update bio_device_shiftlogs set ds_shift = 'H',ds_shift_actual = 'H',ds_status = 'H',ds_shift_begintime = '00:00',ds_shift_endtime = '00:00',ds_begin_duration  = '',ds_end_duration = '' ,ds_sft_hrs = 0  from bio_device_shiftlogs a, emp_dec_holiday_empwise b where emp_decholi_date = ds_date and  emp_decholi_fpcode = ds_fpcode and  emp_decholi_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' "
    paydb.Execute sql
    
    
''    Dim woday As String
    pst_qry = "select * from emp_mas where emp_company in (1,2,3,5) and emp_status = 'A' "
    payrs.Open pst_qry, paydb, 1, 2
    While Not payrs.EOF
         If payrs("emp_holiday") = "SUNDAY" Then
                woday = 1
         ElseIf payrs("emp_holiday") = "MONDAY" Then
                woday = 2
         ElseIf payrs("emp_holiday") = "TUESDAY" Then
                woday = 3
         ElseIf payrs("emp_holiday") = "WEDNESDAY" Then
                woday = 4
         ElseIf payrs("emp_holiday") = "THURSDAY" Then
                woday = 5
         ElseIf payrs("emp_holiday") = "FRIDAY" Then
                woday = 6
         ElseIf payrs("emp_holiday") = "SATURDAY" Then
                woday = 7
         Else
                woday = 0
         End If
         If payrs("emp_cat") = "W" Then
            sql = "update bio_device_shiftlogs set ds_shift_original = 'WOH', ds_status = 'WOH', ds_shift = 'WOH',ds_shift_actual = 'WOH',ds_shift_begintime = '',ds_shift_endtime = '',ds_begin_duration  = '',ds_end_duration = '' where ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and ds_fpcode = " & payrs("emp_fpcode") & " and ds_shift = 'H' and datepart(dw,ds_date) =  " & woday
            sql = "update bio_device_shiftlogs set ds_shift_original = 'WOH', ds_status = 'WOH', ds_shift = 'WOH',ds_shift_begintime = '',ds_shift_endtime = '',ds_begin_duration  = '',ds_end_duration = '' where ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and ds_fpcode = " & payrs("emp_fpcode") & " and ds_shift = 'H' and datepart(dw,ds_date) =  " & woday
            paydb.Execute sql
            sql = "update bio_device_shiftlogs set ds_shift_original = 'WO', ds_status = 'WO', ds_shift = 'WO',ds_shift_actual = 'WO',ds_shift_begintime = '',ds_shift_endtime = '',ds_begin_duration  = '',ds_end_duration = '' where ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and ds_fpcode = " & payrs("emp_fpcode") & " and ds_shift <> 'WOH' and datepart(dw,ds_date) =  " & woday
            sql = "update bio_device_shiftlogs set ds_shift_original = 'WO', ds_status = 'WO', ds_shift = 'WO',ds_shift_begintime = '',ds_shift_endtime = '',ds_begin_duration  = '',ds_end_duration = '' where ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and ds_fpcode = " & payrs("emp_fpcode") & " and ds_shift <> 'WOH' and datepart(dw,ds_date) =  " & woday
            paydb.Execute sql
         
         Else
            sql = "update bio_device_shiftlogs set ds_shift_original = 'WO', ds_status = 'WO', ds_shift = 'WO',ds_shift_actual = 'WO',ds_shift_begintime = '',ds_shift_endtime = '',ds_begin_duration  = '',ds_end_duration = '' where ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and ds_fpcode = " & payrs("emp_fpcode") & " and datepart(dw,ds_date) =  " & woday
            sql = "update bio_device_shiftlogs set ds_shift_original = 'WO', ds_status = 'WO', ds_shift = 'WO',ds_shift_begintime = '',ds_shift_endtime = '',ds_begin_duration  = '',ds_end_duration = '' where ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and ds_fpcode = " & payrs("emp_fpcode") & " and datepart(dw,ds_date) =  " & woday
            paydb.Execute sql
         End If
         
         payrs.MoveNext
    Wend
    payrs.Close


    
''updating voucher payment
    pst_qry = "select * from emp_voupay_mast where emp_company in (1,2,3,5)  and emp_workplace = 'MILL' and emp_cat <> 'M'  "
    payrs.Open pst_qry, paydb, 1, 2
    While Not payrs.EOF
         If payrs("emp_holiday") = "SUNDAY" Then
                woday = 1
         ElseIf payrs("emp_holiday") = "MONDAY" Then
                woday = 2
         ElseIf payrs("emp_holiday") = "TUESDAY" Then
                woday = 3
         ElseIf payrs("emp_holiday") = "WEDNESDAY" Then
                woday = 4
         ElseIf payrs("emp_holiday") = "THURSDAY" Then
                woday = 5
         ElseIf payrs("emp_holiday") = "FRIDAY" Then
                woday = 6
         ElseIf payrs("emp_holiday") = "SATURDAY" Then
                woday = 7
         Else
                woday = 0
         End If
''         If payrs("emp_fpcode") = 1093 Then
''            MsgBox (payrs("emp_fpcode"))
''         End If
         sql = "update bio_device_shiftlogs set ds_shift_original = 'WO',ds_status = 'WO',ds_shift_actual = 'WO',ds_shift = 'WO',ds_shift_begintime = '',ds_shift_endtime = '',ds_begin_duration  = '',ds_end_duration = '',ds_sft_hrs = 0 where ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and ds_fpcode = " & payrs("emp_fpcode") & " and datepart(dw,ds_date) =  " & woday & " "
         
         
         paydb.Execute sql
         payrs.MoveNext
    Wend
    payrs.Close




    Dim emp As Integer
    Dim skip_empcode As String
    skip_empcode = "("
    emp = 0
    i = 0
''    pst_qry = "select * from bio_device_shiftlogs a , bio_empmas b  where  ds_fpcode = bioemp_fpcode and  ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "'  and ds_shift_in > ds_shift_out"
''    payrs.Open pst_qry, paydb, adOpenDynamic, adLockOptimistic
''    While Not payrs.EOF
''          i = i + 1
''''          MsgBox ("Problem in " + Str(payrs!bioemp_fpcode) + " " + payrs!bioemp_name + " in the date of " + Format(payrs!ds_date, "dd/MM/yyyy"))
''          If skip_empcode = "(" Then
''             skip_empcode = skip_empcode + Str(payrs!bioemp_fpcode)
''          Else
''             skip_empcode = skip_empcode + "," + Str(payrs!bioemp_fpcode)
''          End If
''          emp_err(emp) = Str(payrs!bioemp_fpcode)
''          emp = emp + 1
''          payrs.MoveNext
''    Wend
''    payrs.Close
''
''
''    pst_qry = "select * from bio_device_shiftlogs a , bio_empmas b  where  ds_fpcode = bioemp_fpcode and  ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "'  and ds_shift_in = 0  and ds_shift_OUT > 0 "
''    payrs.Open pst_qry, paydb, adOpenDynamic, adLockOptimistic
''    While Not payrs.EOF
''          i = i + 1
''  ''        MsgBox ("Problem in " + Str(payrs!bioemp_fpcode) + " " + payrs!bioemp_name + " in the date of " + Format(payrs!ds_date, "dd/MM/yyyy"))
''          If skip_empcode = "(" Then
''             skip_empcode = skip_empcode + Str(payrs!bioemp_fpcode)
''          Else
''             skip_empcode = skip_empcode + "," + Str(payrs!bioemp_fpcode)
''          End If
''          emp_err(emp) = Str(payrs!bioemp_fpcode)
''          emp = emp + 1
''
''          payrs.MoveNext
''    Wend
''    payrs.Close
''
''''New addition on 30/01/19 start
''
''
''    pst_qry = "select * from bio_device_shiftlogs  a , bio_empmas b    where  ds_fpcode = bioemp_fpcode and   ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "'  and ds_shift_in2 > 0  and ds_shift_out2 = 0"
''    payrs.Open pst_qry, paydb, adOpenDynamic, adLockOptimistic
''
''    While Not payrs.EOF
''          i = i + 1
''''          MsgBox ("Problem in " + Str(payrs!bioemp_fpcode) + " " + payrs!bioemp_name + " in the date of " + Format(payrs!ds_date, "dd/MM/yyyy"))
''          If skip_empcode = "(" Then
''             skip_empcode = skip_empcode + Str(payrs!bioemp_fpcode)
''          Else
''             skip_empcode = skip_empcode + "," + Str(payrs!bioemp_fpcode)
''          End If
''          emp_err(emp) = Str(payrs!bioemp_fpcode)
''          emp = emp + 1
''
''          payrs.MoveNext
''    Wend
''    payrs.Close
''
''    pst_qry = "select * from bio_device_shiftlogs a , bio_empmas b   where  ds_fpcode = bioemp_fpcode and   ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "'  and ds_shift_in3 > 0  and ds_shift_out3 = 0"
''    payrs.Open pst_qry, paydb, adOpenDynamic, adLockOptimistic
''
''    While Not payrs.EOF
''          i = i + 1
''''          MsgBox ("Problem in " + Str(payrs!bioemp_fpcode) + " " + payrs!bioemp_name + " in the date of " + Format(payrs!ds_date, "dd/MM/yyyy"))
''          If skip_empcode = "(" Then
''             skip_empcode = skip_empcode + Str(payrs!bioemp_fpcode)
''          Else
''             skip_empcode = skip_empcode + "," + Str(payrs!bioemp_fpcode)
''          End If
''          emp_err(emp) = Str(payrs!bioemp_fpcode)
''          emp = emp + 1
''
''          payrs.MoveNext
''    Wend
''    payrs.Close
''    skip_empcode = skip_empcode + ")"
''
''''New addition on 30/01/19 end
''
''''    If i > 0 Then Exit Sub
    
100:
    
''    If skip_empcode = "()" Then
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
''    Else
''        sql = "update bio_device_shiftlogs  set ds_sft_hrs1 =  (datediff(minute,ds_shift_in, ds_shift_out)/60)+convert(decimal,(datediff(minute,ds_shift_in, ds_shift_out) -(datediff(minute,ds_shift_in, ds_shift_out)/60 * 60)))/100 where ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and ds_shift_out > ds_shift_in and ds_shift_in > 0  and ds_fpcode   not in " & skip_empcode
''        paydb.Execute sql
''        sql = "update bio_device_shiftlogs  set ds_sft_hrs2 =  (datediff(minute,ds_shift_in2, ds_shift_out2)/60)+convert(decimal,(datediff(minute,ds_shift_in2, ds_shift_out2) -(datediff(minute,ds_shift_in2, ds_shift_out2)/60 * 60)))/100 where ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and ds_shift_out2 > ds_shift_in2 and ds_shift_in2 > 0  and ds_fpcode not in " & skip_empcode
''        paydb.Execute sql
''        sql = "update bio_device_shiftlogs  set ds_sft_hrs3 =  (datediff(minute,ds_shift_in3, ds_shift_out3)/60)+convert(decimal,(datediff(minute,ds_shift_in3, ds_shift_out3) -(datediff(minute,ds_shift_in3, ds_shift_out3)/60 * 60)))/100 where ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and ds_shift_out3 > ds_shift_in3 and ds_shift_in3 > 0  and ds_fpcode not in " & skip_empcode
''        paydb.Execute sql
''        sql = "update bio_device_shiftlogs  set ds_sft_hrs4 =  (datediff(minute,ds_shift_in4, ds_shift_out4)/60)+convert(decimal,(datediff(minute,ds_shift_in4, ds_shift_out4) -(datediff(minute,ds_shift_in4, ds_shift_out4)/60 * 60)))/100 where ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and ds_shift_out4 > ds_shift_in4 and ds_shift_in4 > 0  and ds_fpcode not in " & skip_empcode
''        paydb.Execute sql
''        sql = "update bio_device_shiftlogs  set ds_sft_hrs5 =  (datediff(minute,ds_shift_in5, ds_shift_out5)/60)+convert(decimal,(datediff(minute,ds_shift_in5, ds_shift_out5) -(datediff(minute,ds_shift_in5, ds_shift_out5)/60 * 60)))/100 where ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and ds_shift_out5 > ds_shift_in5 and ds_shift_in5 > 0  and ds_fpcode not in " & skip_empcode
''        paydb.Execute sql
''        sql = "update bio_device_shiftlogs  set ds_sft_hrs6 =  (datediff(minute,ds_shift_in6, ds_shift_out6)/60)+convert(decimal,(datediff(minute,ds_shift_in6, ds_shift_out6) -(datediff(minute,ds_shift_in6, ds_shift_out6)/60 * 60)))/100 where ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and ds_shift_out6 > ds_shift_in6 and ds_shift_in6 > 0  and ds_fpcode not in " & skip_empcode
''        paydb.Execute sql
''    End If

    
''    For i = 0 To emp
''''        If emp_err(i) = 6004 Then
''''           MsgBox (emp_err(i))
''''        End If
''        K = i
''        For idate = st_date To end_date - 1
''''            MsgBox (idate)
''            sql = "update bio_device_shiftlogs  set ds_sft_hrs =  (datediff(minute,ds_shift_in, ds_shift_out)/60)+convert(decimal,(datediff(minute,ds_shift_in, ds_shift_out) -(datediff(minute,ds_shift_in, ds_shift_out)/60 * 60)))/100 where ds_date =  '" & Format(idate, "MM/dd/yyyy") & "' and ds_fpcode = " & emp_err(i) & " and ds_shift_out > ds_shift_in and ds_shift_in > 0 "
''            paydb.Execute sql
''        Next
''    Next
    
''    sql = "update bio_device_shiftlogs  set ds_sft_hrs =  (datediff(minute,ds_shift_in, ds_shift_out)/60)+convert(decimal,(datediff(minute,ds_shift_in, ds_shift_out) -(datediff(minute,ds_shift_in, ds_shift_out)/60 * 60)))/100 where ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "'"
''    paydb.Execute sql
    
''    sql = "update bio_device_shiftlogs set ds_sft_hrs  = 0 WHERE ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(Now - 1, "MM/dd/yyyy") & "'  and ds_shift_in is null and ds_shift_out is null and ds_sft_hrs is null"
''    paydb.Execute sql
''
    
''    Dim firsttime, endtime As Double
''    Dim workedtotmins, workedhrs, workedmins, tothrs As Double
'' ''for updating Permission entries
''    pst_qry = "select * from bio_emp_permissions where empp_date between '" & Format(st_date, "MM/dd/yyyy") & "' and '" & Format(end_date, "MM/dd/yyyy") & "'"
''     payrs.Open pst_qry, paydb, adOpenDynamic, adLockOptimistic
''    While Not payrs.EOF
''          id = payrs!empp_fpcode
''          idate = payrs!empp_date
''
''''          If id = 1099 Then
''''              MsgBox ("wait")
''''          End If
''''
''
''          firsttime = Int(Val(payrs!empp_fromtime)) * 60 + (Val(payrs!empp_fromtime) - Int(Val(payrs!empp_fromtime))) * 100
''          endtime = Int(Val(payrs!empp_totime)) * 60 + (Val(payrs!empp_totime) - Int(Val(payrs!empp_totime))) * 100
''
''           sql = "select * from bio_device_shiftlogs where ds_date = '" & Format(idate, "MM/dd/yyyy") & "' and ds_fpcode = " & id & "  and ds_shift_in is not null "
''           payrs2.Open sql, paydb, 1, 2
''           While Not payrs2.EOF
''                 workedtotmins = DateDiff("n", payrs2!ds_shift_in, payrs2!ds_shift_out) + (endtime - firsttime)
''                 workedhrs = Int(workedtotmins / 60)
''                 workedmins = workedtotmins - (workedhrs * 60)
''                 tothrs = workedhrs + (workedmins / 100)
''                 If tothrs < 0 Then tothrs = 0
''                 payrs2("ds_sft_hrs") = tothrs
''                 payrs2.Update
''                 payrs2.MoveNext
''          Wend
''          payrs2.Close
''          payrs.MoveNext
''    Wend
''    payrs.Close
''

    
    
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
    
    
210:

''New Addition for A SHIFT & B SHIFT 21/07/2020
''Start

''Modified on 20/12/2020 - for Day Closing
    
    sql = "update bio_device_shiftlogs set ds_shift_actual = 'A SHIFT'  from bio_device_shiftlogs WHERE  ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and DATEpart(hour,ds_shift_in) >= 5 and DATEpart(hour,ds_shift_in) <= 6 "
    paydb.Execute sql
    sql = "update bio_device_shiftlogs set ds_shift_actual = 'B SHIFT'  from bio_device_shiftlogs WHERE  ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and DATEpart(hour,ds_shift_in) >= 13 and DATEpart(hour,ds_shift_in) <= 14 "
    paydb.Execute sql
    sql = "update bio_device_shiftlogs set ds_shift_actual = 'C SHIFT'  from bio_device_shiftlogs WHERE  ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and DATEpart(hour,ds_shift_in) >= 21 and DATEpart(hour,ds_shift_in) <= 22 "
    paydb.Execute sql
    
    sql = "update bio_device_shiftlogs set ds_shift_actual = '06.00PM-06.00AM'  from bio_device_shiftlogs WHERE  ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and '" & Format(end_date, "MM/dd/yyyy") & "' and DATEpart(hour,ds_shift_in) >= 17 and DATEpart(hour,ds_shift_in) <= 18"
    paydb.Execute sql
    sql = "update bio_device_shiftlogs set ds_shift_actual = '07.00PM-07.00AM'  from bio_device_shiftlogs WHERE  ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and '" & Format(end_date, "MM/dd/yyyy") & "' and DATEpart(hour,ds_shift_in) >= 18 and DATEpart(hour,ds_shift_in) <= 19"
    paydb.Execute sql
    sql = "update bio_device_shiftlogs set ds_shift_actual = '08.00PM-08.00AM'  from bio_device_shiftlogs WHERE  ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and '" & Format(end_date, "MM/dd/yyyy") & "' and DATEpart(hour,ds_shift_in) >= 19 and DATEpart(hour,ds_shift_in) <= 20"
    paydb.Execute sql
    
    sql = "update bio_device_shiftlogs set ds_shift_actual = 'A SHIFT'  from bio_device_shiftlogs WHERE  ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and DATEpart(hour,ds_shift_in) >= 5 and DATEpart(hour,ds_shift_in) <= 6 and   DATEpart(hour,ds_shift_out) >= 13  and  DATEpart(hour,ds_shift_out) <= 16 "
    paydb.Execute sql
    sql = "update bio_device_shiftlogs set ds_shift_actual = 'B SHIFT'  from bio_device_shiftlogs WHERE  ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and DATEpart(hour,ds_shift_in) >= 13 and DATEpart(hour,ds_shift_in) <= 14 and  DATEpart(hour,ds_shift_out) >= 21  and  DATEpart(hour,ds_shift_out) <= 23 "
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
    sql = "update bio_device_shiftlogs set ds_status = '½P'  from bio_device_shiftlogs , bio_empmas   WHERE  ds_fpcode = bioemp_fpcode and ds_sft_hrs >3  and  ds_sft_hrs < 7.49 and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and ds_shift <> 'WO' and ds_shift <> 'WOH' and ds_shift_actual in ('A SHIFT','B SHIFT','C SHIFT')"
    paydb.Execute sql
    sql = "update bio_device_shiftlogs set ds_status = 'WO½P'  from bio_device_shiftlogs , bio_empmas   WHERE  ds_fpcode = bioemp_fpcode and ds_sft_hrs >3  and  ds_sft_hrs < 7.49 and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and ds_shift = 'WO' and ds_shift <> 'WOH' and ds_shift_actual in ('A SHIFT','B SHIFT','C SHIFT')"
    paydb.Execute sql
    
''FOR STAFF - 1/2 DAY PRESENT - for SINGLE IN / OUT punches
    sql = "update bio_device_shiftlogs set ds_status = '½P'  from bio_device_shiftlogs , bio_empmas   WHERE  ds_fpcode = bioemp_fpcode and bioemp_team = 'STAFF' and   ds_sft_hrs >3  and  ds_sft_hrs < 8.19 and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and ds_shift <> 'WO' and ds_shift <> 'WOH' and ds_no_of_punches = 1 and ds_shift_actual not in ('A SHIFT','B SHIFT','C SHIFT')"
    paydb.Execute sql
    sql = "update bio_device_shiftlogs set ds_status = 'WO½P'  from bio_device_shiftlogs , bio_empmas   WHERE  ds_fpcode = bioemp_fpcode and bioemp_team = 'STAFF' and  ds_sft_hrs >3  and  ds_sft_hrs < 8.19 and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and ds_shift = 'WO' and ds_shift <> 'WOH' and ds_no_of_punches = 1 and ds_shift_actual not in ('A SHIFT','B SHIFT','C SHIFT')"
    paydb.Execute sql
    
''FOR UNSHIFT
    sql = "update bio_device_shiftlogs set ds_status = '½P'  from bio_device_shiftlogs , bio_empmas   WHERE  ds_fpcode = bioemp_fpcode and bioemp_team = 'STAFF' and   ds_sft_hrs >3  and  ds_sft_hrs < 7.49 and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and ds_shift <> 'WO' and ds_shift <> 'WOH' and ds_no_of_punches = 1 and ds_shift_actual  in ('Unshift')"
    sql = "update bio_device_shiftlogs set ds_status = '½P'  from bio_device_shiftlogs , bio_empmas   WHERE  ds_fpcode = bioemp_fpcode and   ds_sft_hrs >3  and  ds_sft_hrs < 7.49 and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and ds_shift <> 'WO' and ds_shift <> 'WOH' and ds_no_of_punches = 1 and ds_shift_actual  in ('Unshift')"
    
    paydb.Execute sql
    
    sql = "update bio_device_shiftlogs set ds_status = 'P'  from bio_device_shiftlogs , bio_empmas   WHERE  ds_fpcode = bioemp_fpcode and bioemp_team = 'STAFF' and   ds_sft_hrs > 7.48  and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and ds_shift <> 'WO' and ds_shift <> 'WOH' and ds_no_of_punches = 1 and ds_shift_actual  in ('Unshift') "
    sql = "update bio_device_shiftlogs set ds_status = 'P'  from bio_device_shiftlogs , bio_empmas   WHERE  ds_fpcode = bioemp_fpcode and   ds_sft_hrs > 7.48  and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and ds_shift <> 'WO' and ds_shift <> 'WOH' and ds_no_of_punches = 1 and ds_shift_actual  in ('Unshift') "
    paydb.Execute sql
    
''FOR STAFF - 1/2 DAY PRESENT - for DOUBLE IN / OUT punches
    sql = "update bio_device_shiftlogs set ds_status = '½P'  from bio_device_shiftlogs , bio_empmas   WHERE  ds_fpcode = bioemp_fpcode and bioemp_team = 'STAFF' and   ds_sft_hrs >3  and  ds_sft_hrs < 7.49 and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and ds_shift <> 'WO' and ds_shift <> 'WOH' and ds_no_of_punches > 1 and ds_shift_actual not in ('A SHIFT','B SHIFT','C SHIFT')"
    paydb.Execute sql
    sql = "update bio_device_shiftlogs set ds_status = 'WO½P'  from bio_device_shiftlogs , bio_empmas   WHERE  ds_fpcode = bioemp_fpcode and bioemp_team = 'STAFF' and  ds_sft_hrs >3  and  ds_sft_hrs < 7.49 and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and ds_shift = 'WO' and ds_shift <> 'WOH' and ds_no_of_punches > 1 and ds_shift_actual not in ('A SHIFT','B SHIFT','C SHIFT')"
    paydb.Execute sql
    
    
''FOR DECLARE HOLIDAY - STAFF - 1/2 DAY PRESENT
    sql = "update bio_device_shiftlogs set ds_status = '½HP'  from bio_device_shiftlogs , bio_empmas   WHERE  ds_fpcode = bioemp_fpcode and bioemp_team = 'STAFF' and  ds_sft_hrs > 4 and  ds_sft_hrs < 7.40 and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and  ds_shift = 'H'  "
    paydb.Execute sql
    sql = "update bio_device_shiftlogs set ds_status = 'HP' from bio_device_shiftlogs , bio_empmas    WHERE ds_fpcode = bioemp_fpcode and bioemp_team = 'STAFF' and ds_sft_hrs >= 7.40 and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and  ds_shift = 'H' "
    paydb.Execute sql
    
    
''FOR WORKER - 1/2 DAY PRESENT
    sql = "update bio_device_shiftlogs set ds_status = '½P'  from bio_device_shiftlogs , bio_empmas   WHERE  ds_fpcode = bioemp_fpcode and bioemp_team <> 'STAFF' and   ds_sft_hrs >3  and  ds_sft_hrs < 7.50 and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and ds_shift <> 'WO' and ds_shift <> 'WOH'  "
    paydb.Execute sql
    sql = "update bio_device_shiftlogs set ds_status = 'WO½P'  from bio_device_shiftlogs , bio_empmas   WHERE  ds_fpcode = bioemp_fpcode and bioemp_team <> 'STAFF' and  ds_sft_hrs >3  and  ds_sft_hrs < 7.50 and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and ds_shift = 'WO' and ds_shift <> 'WOH'  "
    paydb.Execute sql
    sql = "update bio_device_shiftlogs set ds_status = '½HP'  from bio_device_shiftlogs , bio_empmas   WHERE  ds_fpcode = bioemp_fpcode and bioemp_team <> 'STAFF' and  ds_sft_hrs > 4 and  ds_sft_hrs < 7.50 and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and  ds_shift = 'H'"
    paydb.Execute sql
    sql = "update bio_device_shiftlogs set ds_status = 'HP' from bio_device_shiftlogs , bio_empmas    WHERE ds_fpcode = bioemp_fpcode and bioemp_team <> 'STAFF' and ds_sft_hrs >= 7.50 and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and  ds_shift = 'H'"
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
    sql = "update bio_device_shiftlogs set ds_status = 'P' from bio_device_shiftlogs , bio_empmas  WHERE ds_fpcode = bioemp_fpcode and ds_sft_hrs >= 7.49 and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "'  and  ds_shift <> 'H'  and  ds_shift <> 'WOH'  and ds_shift_actual in ('A SHIFT','B SHIFT','C SHIFT')"
    paydb.Execute sql
    sql = "update bio_device_shiftlogs set ds_status = 'WOP' from bio_device_shiftlogs , bio_empmas  WHERE ds_fpcode = bioemp_fpcode and  ds_sft_hrs >= 7.49 and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and  ds_shift = 'WO' and ds_status = 'P'  and ds_shift_actual in ('A SHIFT','B SHIFT','C SHIFT')"
    paydb.Execute sql
    

    '' for  STAFF
    sql = "update bio_device_shiftlogs set ds_status = 'P' from bio_device_shiftlogs , bio_empmas  WHERE ds_fpcode = bioemp_fpcode and bioemp_team = 'STAFF' and ds_sft_hrs >= 8.19 and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "'  and  ds_shift <> 'H'  and  ds_shift <> 'WOH' and ds_no_of_punches = 1 and ds_shift_actual not in ('A SHIFT','B SHIFT','C SHIFT')"
    paydb.Execute sql
    sql = "update bio_device_shiftlogs set ds_status = 'WOP' from bio_device_shiftlogs , bio_empmas  WHERE ds_fpcode = bioemp_fpcode and bioemp_team = 'STAFF' and  ds_sft_hrs >= 8.19 and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and  ds_shift = 'WO' and ds_status = 'P' and ds_no_of_punches = 1   and ds_shift_actual not in ('A SHIFT','B SHIFT','C SHIFT')"
    paydb.Execute sql
    
    sql = "update bio_device_shiftlogs set ds_status = 'WOHP' from bio_device_shiftlogs , bio_empmas  WHERE ds_fpcode = bioemp_fpcode and bioemp_team = 'STAFF' and ds_sft_hrs >= 8.19 and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and  ds_status = 'WOH'  and ds_no_of_punches = 1 and ds_shift_actual not in ('A SHIFT','B SHIFT','C SHIFT')"
    paydb.Execute sql
    sql = "update bio_device_shiftlogs set ds_status = 'P' from bio_device_shiftlogs , bio_empmas WHERE ds_fpcode = bioemp_fpcode and bioemp_team = 'STAFF' and ds_sft_hrs >= 8.19 and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "'  and  ds_shift <> 'H' and  ds_shift <> 'WOH'  and  ds_shift <> 'WO' and ds_no_of_punches = 1  and ds_shift_actual not in ('A SHIFT','B SHIFT','C SHIFT')"
    paydb.Execute sql
    
    sql = "update bio_device_shiftlogs set ds_status = 'P' from bio_device_shiftlogs , bio_empmas  WHERE ds_fpcode = bioemp_fpcode and bioemp_team = 'STAFF' and ds_sft_hrs >= 7.49 and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "'  and  ds_shift <> 'H'  and  ds_shift <> 'WOH' and ds_no_of_punches > 1 and ds_shift_actual not in ('A SHIFT','B SHIFT','C SHIFT')"
    paydb.Execute sql
    sql = "update bio_device_shiftlogs set ds_status = 'WOP' from bio_device_shiftlogs , bio_empmas  WHERE ds_fpcode = bioemp_fpcode and bioemp_team = 'STAFF' and  ds_sft_hrs >= 7.49 and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and  ds_shift = 'WO' and ds_status = 'P' and ds_no_of_punches > 1   and ds_shift_actual not in ('A SHIFT','B SHIFT','C SHIFT')"
    paydb.Execute sql
    
    sql = "update bio_device_shiftlogs set ds_status = 'WOHP' from bio_device_shiftlogs , bio_empmas  WHERE ds_fpcode = bioemp_fpcode and bioemp_team = 'STAFF' and ds_sft_hrs >= 7.49 and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and  ds_status = 'WOH'  and ds_no_of_punches > 1 and ds_shift_actual not in ('A SHIFT','B SHIFT','C SHIFT')"
    paydb.Execute sql
    sql = "update bio_device_shiftlogs set ds_status = 'P' from bio_device_shiftlogs , bio_empmas WHERE ds_fpcode = bioemp_fpcode and bioemp_team = 'STAFF' and ds_sft_hrs >= 7.49 and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "'  and  ds_shift <> 'H' and  ds_shift <> 'WOH'  and  ds_shift <> 'WO' and ds_no_of_punches > 1  and ds_shift_actual not in ('A SHIFT','B SHIFT','C SHIFT')"
    paydb.Execute sql
    
    
    
    sql = "update bio_device_shiftlogs set ds_status = 'P' from bio_device_shiftlogs , bio_empmas  WHERE ds_fpcode = bioemp_fpcode and bioemp_team <> 'STAFF' and ds_sft_hrs >= 7.50 and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "'  and  ds_shift <> 'H'  and  ds_shift <> 'WOH' and ds_shift_actual not in ('A SHIFT','B SHIFT','C SHIFT')"
    paydb.Execute sql
    sql = "update bio_device_shiftlogs set ds_status = 'WOP' from bio_device_shiftlogs , bio_empmas  WHERE ds_fpcode = bioemp_fpcode and bioemp_team <> 'STAFF' and  ds_sft_hrs >= 7.50 and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and  ds_shift = 'WO' and ds_status = 'P' and ds_shift_actual not in ('A SHIFT','B SHIFT','C SHIFT')"
    paydb.Execute sql
    sql = "update bio_device_shiftlogs set ds_status = 'WOHP' from bio_device_shiftlogs , bio_empmas  WHERE ds_fpcode = bioemp_fpcode and bioemp_team <> 'STAFF' and ds_sft_hrs >= 7.50 and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and  ds_status = 'WOH' and ds_shift_actual not in ('A SHIFT','B SHIFT','C SHIFT')"
    paydb.Execute sql
    sql = "update bio_device_shiftlogs set ds_status = 'P' from bio_device_shiftlogs , bio_empmas WHERE ds_fpcode = bioemp_fpcode and bioemp_team <> 'STAFF' and ds_sft_hrs >= 7.50 and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "'  and  ds_shift <> 'H' and  ds_shift <> 'WOH'  and  ds_shift <> 'WO' and ds_shift_actual not in ('A SHIFT','B SHIFT','C SHIFT') "
    paydb.Execute sql
    
    
''FOR SECURITY GUARD
    sql = "update bio_device_shiftlogs set ds_status = '½P'  from bio_device_shiftlogs , bio_empmas   WHERE  ds_fpcode = bioemp_fpcode and bioemp_team = 'SECURITY GUARD' and   ds_sft_hrs >4  and  ds_sft_hrs < 11.30 and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "'"
    paydb.Execute sql
    sql = "update bio_device_shiftlogs set ds_status = 'P'  from bio_device_shiftlogs , bio_empmas   WHERE  ds_fpcode = bioemp_fpcode and bioemp_team = 'SECURITY GUARD' and   ds_sft_hrs >=11.30  and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "'"
    paydb.Execute sql
    
'' for OD Assignment
    sql = "update bio_device_shiftlogs set ds_status = 'P(OD)' from bio_device_shiftlogs , bio_empmas  WHERE ds_fpcode = bioemp_fpcode and  ds_sft_hrs >= 7.49 and  ds_od_hrs >= 4   and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' "
    paydb.Execute sql
    sql = "update bio_device_shiftlogs set ds_status = 'WOP(OD)' from bio_device_shiftlogs , bio_empmas  WHERE ds_fpcode = bioemp_fpcode and ds_sft_hrs >= 7.49 and  ds_od_hrs >= 4 and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and  ds_shift = 'WO'"
    paydb.Execute sql
    
    ''FOR HOUSE KEEPING - MINIMUM 1HRS WORKING
    sql = "update bio_device_shiftlogs set ds_status = 'P'  from bio_device_shiftlogs , bio_empmas   WHERE  ds_fpcode in (10071,10039,10103,10102,10019,10015,10023) and ds_fpcode = bioemp_fpcode and   ds_sft_hrs >=1  and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "'"
    paydb.Execute sql
    
    
    ''FOR MESS - MINIMUM 6 HRS WORKING (SIX HOURS)
    sql = "update bio_device_shiftlogs set ds_status = 'P'  from bio_device_shiftlogs , bio_empmas   WHERE  ds_fpcode in (10222,10221) and ds_fpcode = bioemp_fpcode and   ds_sft_hrs >=6  and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "'"
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
                  leavetype = IIf(payrs!emp_leave_period = "F", "EL", "½EL")
               ElseIf payrs!emp_leave_type = "PL" Then
                  leavetype = IIf(payrs!emp_leave_period = "F", "PL", "½PL")
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
          If leavetype = "½PL" And payrs!ds_status = "½P" Then
             sql = "update bio_device_shiftlogs set ds_status = '½P½PL' where ds_fpcode =  '" & id & "' and ds_date =  '" & Format(payrs!emp_leave_date, "MM/dd/yyyy") & "'"
          ElseIf leavetype = "½EL" And payrs!ds_status = "½P" Then
             sql = "update bio_device_shiftlogs set ds_status = '½P½EL' where ds_fpcode =  '" & id & "' and ds_date =  '" & Format(payrs!emp_leave_date, "MM/dd/yyyy") & "'"
             '''''''' modified by devaraj
''          ElseIf leavetype = "½EL½A" And payrs!ds_status = "½PL" Then
''             sql = "update bio_device_shiftlogs set ds_status = '½EL½PL' where ds_fpcode =  '" & id & "' and ds_date =  '" & Format(payrs!emp_leave_date, "MM/dd/yyyy") & "'"
''
''          ElseIf leavetype = "½EL½A" And payrs!ds_status = "½EL" Then
''             sql = "update bio_device_shiftlogs set ds_status = 'EL' where ds_fpcode =  '" & id & "' and ds_date =  '" & Format(payrs!emp_leave_date, "MM/dd/yyyy") & "'"
''          ElseIf leavetype = "½PL½A" And payrs!ds_status = "½EL" Then
''             sql = "update bio_device_shiftlogs set ds_status = '½PL½EL' where ds_fpcode =  '" & id & "' and ds_date =  '" & Format(payrs!emp_leave_date, "MM/dd/yyyy") & "'"
''        ElseIf leavetype = "½PL½A" And payrs!ds_status = "½PL" Then
''             sql = "update bio_device_shiftlogs set ds_status = 'PL' where ds_fpcode =  '" & id & "' and ds_date =  '" & Format(payrs!emp_leave_date, "MM/dd/yyyy") & "'"
            '''''''''''''
            
          ElseIf payrs!ds_status = "½P" And leavetype = "½C.H" Then
             sql = "update bio_device_shiftlogs set ds_status = '½P½C.H' where ds_fpcode =  '" & id & "' and ds_date =  '" & Format(payrs!emp_leave_date, "MM/dd/yyyy") & "'"
          ElseIf payrs!ds_status = "½C.H" And leavetype = "½C.H" Then
             sql = "update bio_device_shiftlogs set ds_status = 'C.H' where ds_fpcode =  '" & id & "' and ds_date =  '" & Format(payrs!emp_leave_date, "MM/dd/yyyy") & "'"
          ElseIf payrs!ds_status = "½C.H" And (leavetype = "½EL" Or leavetype = "½CL") Then
             sql = "update bio_device_shiftlogs set ds_status = '½C.H½EL' where ds_fpcode =  '" & id & "' and ds_date =  '" & Format(payrs!emp_leave_date, "MM/dd/yyyy") & "'"
          ElseIf payrs!ds_status = "½C.H" And leavetype = "½PL" Then
             sql = "update bio_device_shiftlogs set ds_status = '½C.H½PL' where ds_fpcode =  '" & id & "' and ds_date =  '" & Format(payrs!emp_leave_date, "MM/dd/yyyy") & "'"
          ElseIf payrs!ds_status = "P" Then
             sql = "update bio_device_shiftlogs set ds_status = 'P' where ds_fpcode =  '" & id & "' and ds_date =  '" & Format(payrs!emp_leave_date, "MM/dd/yyyy") & "'"
          ElseIf payrs!ds_status <> "A" Then
             sql = "update bio_device_shiftlogs set ds_status = ds_status+'" & leavetype & "' where ds_fpcode =  '" & id & "' and ds_date =  '" & Format(payrs!emp_leave_date, "MM/dd/yyyy") & "'"
          ElseIf payrs!ds_status = "A" And leavetype = "½EL" Then
             sql = "update bio_device_shiftlogs set ds_status = '½EL½A' where ds_fpcode =  '" & id & "' and ds_date =  '" & Format(payrs!emp_leave_date, "MM/dd/yyyy") & "'"
          ElseIf payrs!ds_status = "A" And leavetype = "½P" Then
             sql = "update bio_device_shiftlogs set ds_status = '½P½A' where ds_fpcode =  '" & id & "' and ds_date =  '" & Format(payrs!emp_leave_date, "MM/dd/yyyy") & "'"
          ElseIf payrs!ds_status = "A" And leavetype = "½PL" Then
             sql = "update bio_device_shiftlogs set ds_status = '½PL½A' where ds_fpcode =  '" & id & "' and ds_date =  '" & Format(payrs!emp_leave_date, "MM/dd/yyyy") & "'"
          Else
             sql = "update bio_device_shiftlogs set ds_status = '" & leavetype & "' where ds_fpcode =  '" & id & "' and ds_date =  '" & Format(payrs!emp_leave_date, "MM/dd/yyyy") & "'"
          End If
          paydb2.Execute sql
          payrs.MoveNext
    Wend
    payrs.Close
 
 ''for updating OD entries
''    pst_qry = "select * from bio_emp_oddetails where empod_date between '" & Format(st_date, "MM/dd/yyyy") & "' and '" & Format(end_date, "MM/dd/yyyy") & "'"
''    payrs.Open pst_qry, paydb, adOpenDynamic, adLockOptimistic
''    While Not payrs.EOF
''          id = payrs!empod_fpcode
''          leave = "P(OD)"
''          sql = "update bio_device_shiftlogs set ds_status = 'WOP(OD)' where ds_fpcode =  '" & id & "' and ds_date =  '" & Format(payrs!empod_date, "MM/dd/yyyy") & "' and ds_shift = 'WO'"
''          paydb.Execute sql
''          sql = "update bio_device_shiftlogs set ds_status = 'P(OD)' where ds_fpcode =  '" & id & "' and ds_date =  '" & Format(payrs!empod_date, "MM/dd/yyyy") & "' and ds_shift <> 'WO'"
''          paydb.Execute sql
''
''          payrs.MoveNext
''    Wend
''    payrs.Close
' ''for updating OD entries
350:
''    Dim fday_od As Integer
''    pst_qry = "select * from bio_emp_oddetails where empod_date between '" & Format(st_date, "MM/dd/yyyy") & "' and '" & Format(end_date, "MM/dd/yyyy") & "'"
''    payrs.Open pst_qry, paydb, adOpenDynamic, adLockOptimistic
''    While Not payrs.EOF
''          id = payrs!empod_fpcode
''''
''''          If id = 6011 Then
''''             MsgBox ("Test")
''''          End If
''
''          idate = payrs!empod_date
''
''          firsttime = Int(Val(payrs!empod_fromtime)) * 60 + (Val(payrs!empod_fromtime) - Int(Val(payrs!empod_fromtime))) * 100
''          endtime = Int(Val(payrs!empod_totime)) * 60 + (Val(payrs!empod_totime) - Int(Val(payrs!empod_totime))) * 100
''
''          If endtime - firsttime >= 540 Then
''             fday_od = 1
''          Else
''             fday_od = 0
''          End If
''
''    ''      leave = "P(OD)"
''
''           sql = "select * from bio_device_shiftlogs where ds_date = '" & Format(idate, "MM/dd/yyyy") & "' and ds_fpcode = " & id & "  and ds_shift_in is not null "
''           payrs2.Open sql, paydb, 1, 2
''           While Not payrs2.EOF
''                 If payrs2!ds_shift_in = "01/01/1900" Or payrs2!ds_shift_out = "01/01/1900" Then
''                    workedtotmins = 0
''                    workedhrs = 0
''                    workedmins = 0
''                 Else
''                    workedtotmins = DateDiff("n", payrs2!ds_shift_in, payrs2!ds_shift_out) + (endtime - firsttime)
''                    workedhrs = Int(workedtotmins / 60)
''                    workedmins = workedtotmins - (workedhrs * 60)
''                 End If
''                 tothrs = workedhrs + (workedmins / 100)
''''                 If payrs2("ds_shift") = "WO" Then
''''                    payrs2("ds_status") = "WOP(OD)"
''''                 Else
''''                    payrs2("ds_status") = "P(OD)"
''''                 End If
''                 If payrs2("ds_shift") = "WO" Then
''                    payrs2("ds_status") = "WOP(OD)"
''                 ElseIf payrs2("ds_shift") = "H" Then
''                    payrs2("ds_status") = "HP(OD)"
''                 ElseIf fday_od = 1 Then
''                    payrs2("ds_status") = "P(OD)"
''                 Else
''                    If tothrs > 8 Then
''                       payrs2("ds_status") = "P(OD)"
''
''                       Else
''                          If tothrs >= 3 And payrs2("ds_status") = "½PL½A" Then
''                             payrs2("ds_status") = "½PL½OD"
''                          ElseIf tothrs >= 3 And payrs2("ds_status") = "½EL½A" Then
''                             payrs2("ds_status") = "½EL½OD"
''                          ElseIf tothrs >= 3 And payrs2("ds_status") = "A" Then
''                             payrs2("ds_status") = "½A½OD"
''                          Else
''                             payrs2("ds_status") = "A"
''                          End If
''                       End If
''                    End If
''
''
''                 If tothrs < 0 Then tothrs = 0
''                 payrs2("ds_sft_hrs") = tothrs
''                 payrs2.Update
''                 payrs2.MoveNext
''          Wend
''          payrs2.Close
''          payrs.MoveNext
''
''
''
''''          sql = "update bio_device_shiftlogs set ds_status = 'WOP(OD)' where ds_fpcode =  '" & id & "' and ds_date =  '" & Format(payrs!empod_date, "MM/dd/yyyy") & "' and ds_shift = 'WO'"
''''          paydb.Execute sql
''''          sql = "update bio_device_shiftlogs set ds_status = 'P(OD)' where ds_fpcode =  '" & id & "' and ds_date =  '" & Format(payrs!empod_date, "MM/dd/yyyy") & "' and ds_shift <> 'WO'"
''''          paydb.Execute sql
''''
''''          payrs.MoveNext
''    Wend
''    payrs.Close '
''
 
 
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
          ElseIf payrs!ds_status = "½PL½A" Then
             sql = "update bio_device_shiftlogs set ds_status = '½PL½C.H' where ds_fpcode =  '" & id & "' and ds_date =  '" & Format(payrs!empch_ch_date, "MM/dd/yyyy") & "'"
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
 
 
''for updating holiday present eligibility
''Start


  sql = "update bio_device_shiftlogs set ds_status = 'HPE' from bio_device_shiftlogs , bio_empmas ,emp_mas where bioemp_fpcode = emp_fpcode and bioemp_fpcode = ds_fpcode and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and emp_cat = 'W' and ds_status = 'HP' and emp_dh_wages_yn = 'Y'  and emp_status = 'A'"
  paydb.Execute sql


  sql = "update bio_device_shiftlogs set ds_status = 'HPE' from bio_device_shiftlogs a, bio_empdh_eligible b where empdh_date = ds_date and  empdh_fpcode = ds_fpcode and ds_Status = 'HP' and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "'"
  paydb.Execute sql
  
  
  sql = "update bio_device_shiftlogs set ds_status = 'HPE' from bio_device_shiftlogs , bio_empmas ,emp_mas where bioemp_fpcode = emp_fpcode and bioemp_fpcode = ds_fpcode and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and emp_cat = 'W' and ds_status = 'P(OD)'  and ds_shift = 'H' and emp_dh_wages_yn = 'Y' and emp_status = 'A'"
  paydb.Execute sql
'''MODIFIED BY DEVA'''''''
  sql = "update bio_device_shiftlogs set ds_status = 'HPE' from bio_device_shiftlogs a, bio_empdh_eligible b where empdh_date = ds_date and  empdh_fpcode = ds_fpcode and ds_Status = 'OD' and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "'"
  paydb.Execute sql
  ''''''''''''

  sql = "update bio_device_shiftlogs set ds_status = 'HPE' from bio_device_shiftlogs a, bio_empdh_eligible b where empdh_date = ds_date and  empdh_fpcode = ds_fpcode and ds_Status = 'P(OD)' and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "'"
  paydb.Execute sql
  
  
  sql = "update bio_device_shiftlogs set ds_status = 'WOHPE' from bio_device_shiftlogs a, bio_empdh_eligible b where empdh_date = ds_date and  empdh_fpcode = ds_fpcode and ds_Status = 'WOHP' and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "'"
  paydb.Execute sql

  sql = "update bio_device_shiftlogs set ds_status = 'WOHPE' from bio_device_shiftlogs , bio_empmas ,emp_mas where bioemp_fpcode = emp_fpcode and bioemp_fpcode = ds_fpcode and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and emp_cat = 'W' and ds_status = 'WOHP'   and emp_status = 'A'"
  paydb.Execute sql
  
''End
 
''for 1/2CH and 1/2 ABS
  sql = "update bio_device_shiftlogs set ds_status = '½C.H½A' from bio_device_shiftlogs where ds_Status = '½C.H' and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and ds_sft_hrs = 0"
  paydb.Execute sql
 
 ''for 1/2CH and 1/2 PRESENT
  sql = "update bio_device_shiftlogs set ds_status = '½C.H½P' from bio_device_shiftlogs where ds_Status = '½C.H' and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and ds_sft_hrs > 3.4"
  paydb.Execute sql
 
 
 
'' FOR SATURDAY ATTENDANCE REMOVE

    pst_qry = "update bio_device_shiftlogs set ds_status = 'C' from bio_device_shiftlogs ,bio_emp_saturday where ds_fpcode = empsat_fpcode and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and '" & Format(end_date, "MM/dd/yyyy") & "' and datepart(dw,ds_date) =  7 and ds_status <> 'H'"
    pst_qry = "update bio_device_shiftlogs set ds_status = 'C' from bio_device_shiftlogs ,bio_emp_saturday where ds_fpcode = empsat_fpcode and ds_date =empsat_date"
    
    paydb.Execute pst_qry
 
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
 
500:

    Dim dayfind, dayfind_intime, dayfind_outtime As String
    Dim present, absent, hop, wop, cl, sl, h, ch, layoff, wo, pl, hope, el, woh, ml As Single
 
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
                ElseIf payrs.Fields(dayfind) = "PL" Or payrs.Fields(dayfind) = "PLP" Then
                    pl = pl + 1
                ElseIf payrs.Fields(dayfind) = "½PL" Then
                    pl = pl + 0.5
                    absent = absent + 0.5
                ElseIf payrs.Fields(dayfind) = "½PLP" Then
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
                    hop = hop + 0
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
                    h = h + 1
                ElseIf payrs.Fields(dayfind) = "WOHP" Then
                    woh = woh + 1
                    h = h + 1
                    wop = wop + 1
                ElseIf payrs.Fields(dayfind) = "WOHPE" Then
                    woh = woh + 1
                    h = h + 1
                    wop = wop + 1
                    hope = hope + 1


''                ElseIf payrs.Fields(dayfind) = "½C.H" Then
''                    ch = ch + 0.5
''                    present = present + 0.5
                
                ElseIf payrs.Fields(dayfind) = "½EL½C.H" Then
                    ch = ch + 0.5
                    el = el + 0.5
                
                ElseIf payrs.Fields(dayfind) = "½EL½PL" Then
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
                ElseIf payrs.Fields(dayfind) = "½P½PL" Then
                    pl = pl + 0.5
                    present = present + 0.5
                ElseIf payrs.Fields(dayfind) = "½PL½C.H" Then
                    pl = pl + 0.5
                    ch = ch + 0.5
                ElseIf payrs.Fields(dayfind) = "½P½A" Or payrs.Fields(dayfind) = "½A½P" Then
                    present = present + 0.5
                    absent = absent + 0.5
                ElseIf payrs.Fields(dayfind) = "½EL½A" Or payrs.Fields(dayfind) = "½A½EL" Then
                    el = el + 0.5
                    absent = absent + 0.5
                ElseIf payrs.Fields(dayfind) = "½PL½A" Or payrs.Fields(dayfind) = "½A½PL" Then
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
                ElseIf payrs.Fields(dayfind) = "½PL½OD" Then
                    pl = pl + 0.5
                    present = present + 0.5
                ElseIf payrs.Fields(dayfind) = "½EL½OD" Then
                    el = el + 0.5
                    present = present + 0.5
                ElseIf payrs.Fields(dayfind) = "½ML" Then
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
            payrs("a_wo") = wo
            payrs("a_woh") = woh
            
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
    pst_qry = "update bio_attendlogs set a_hop  = 0 where a_year = " & Val(cmb_year.Text) & "  and a_month = " & cmb_month.ItemData(cmb_month.ListIndex) & " and a_hop < 0 "
    paydb2.Execute pst_qry
    
'end
 
 
    Dim entime2 As Date
    endtime2 = TimeValue(Now)
 
    MsgBox ("Updation completed... Process Start by " + Str(stime) + " ...  end by " + Str(endtime2))
    
 
    pst_qry = "select * from bio_device_shiftlogs a , bio_empmas b  where  ds_fpcode = bioemp_fpcode and  ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "'  and ds_shift_in > ds_shift_out and ds_sft_hrs <2 and ds_date < '" & Format(end_date, "MM/dd/YYYY") & "'"
    payrs.Open pst_qry, paydb, adOpenDynamic, adLockOptimistic
    While Not payrs.EOF
          If Format(payrs!ds_date, "MM/dd/yyyy") <> Format(Now, "MM/dd/yyyy") Then
             MsgBox ("Problem in " + Str(payrs!bioemp_fpcode) + " " + payrs!bioemp_name + " in the date of " + Format(payrs!ds_date, "dd/MM/yyyy"))
          End If
          payrs.MoveNext
    Wend
    payrs.Close


    pst_qry = "select * from bio_device_shiftlogs a , bio_empmas b  where  ds_fpcode = bioemp_fpcode and  ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "'  and ds_shift_in = 0  and ds_shift_OUT > 0 and ds_sft_hrs <2  and ds_date < '" & Format(end_date.Value, "MM/dd/YYYY") & "'"
    payrs.Open pst_qry, paydb, adOpenDynamic, adLockOptimistic
    While Not payrs.EOF
          MsgBox ("Problem in " + Str(payrs!bioemp_fpcode) + " " + payrs!bioemp_name + " in the date of " + Format(payrs!ds_date, "dd/MM/yyyy"))
          payrs.MoveNext
    Wend
    payrs.Close

    
    pst_qry = "select * from bio_device_shiftlogs  a , bio_empmas b    where  ds_fpcode = bioemp_fpcode and   ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "'  and ds_shift_in2 > 0  and ds_shift_out2 = 0 and ds_sft_hrs <2 and  ds_date < '" & Format(end_date.Value, "MM/dd/YYYY") & "'"
    payrs.Open pst_qry, paydb, adOpenDynamic, adLockOptimistic

    While Not payrs.EOF
          MsgBox ("Problem in " + Str(payrs!bioemp_fpcode) + " " + payrs!bioemp_name + " in the date of " + Format(payrs!ds_date, "dd/MM/yyyy"))
          payrs.MoveNext
    Wend
    payrs.Close
    
    
    pst_qry = "select * from bio_device_shiftlogs  a , bio_empmas b    where  ds_fpcode = bioemp_fpcode and   ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "'  and ds_shift_in2 > 0  and ds_shift_out2 = 0 and ds_sft_hrs >24 and  ds_date < '" & Format(end_date.Value, "MM/dd/YYYY") & "'"
    payrs.Open pst_qry, paydb, adOpenDynamic, adLockOptimistic

    While Not payrs.EOF
          MsgBox ("Problem in " + Str(payrs!bioemp_fpcode) + " " + payrs!bioemp_name + " in the date of " + Format(payrs!ds_date, "dd/MM/yyyy"))
          payrs.MoveNext
    Wend
    payrs.Close
    
    
    
    pst_qry = "select * from bio_device_shiftlogs a , bio_empmas b   where  ds_fpcode = bioemp_fpcode and   ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "'  and ds_shift_in3 > 0  and ds_shift_out3 = 0 and ds_sft_hrs < 3  and ds_date < '" & Format(end_date.Value, "MM/dd/YYYY") & "'"
    payrs.Open pst_qry, paydb, adOpenDynamic, adLockOptimistic

    While Not payrs.EOF
          i = i + 1
          MsgBox ("Problem in " + Str(payrs!bioemp_fpcode) + " " + payrs!bioemp_name + " in the date of " + Format(payrs!ds_date, "dd/MM/yyyy"))
          payrs.MoveNext
    Wend
    payrs.Close
    
    
    pst_qry = "select * from bio_device_shiftlogs a , bio_empmas b   where  ds_fpcode = bioemp_fpcode and   ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "'  and ds_shift_in3 > 0  and ds_shift_out3 = 0 and ds_sft_hrs > 24  and ds_date < '" & Format(end_date.Value, "MM/dd/YYYY") & "'"
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
    
    Dim id, fcode As Integer
    Dim dlogdate As Date
    
    Dim dev_log(100) As Long
    
    Dim log_details As String
    
    
    Dim stime, etime As Date
    
    stime = TimeValue(Now)
    
    
    Dim sft, sft_bt, sft_et, sft_begin_dur, sft_end_dur As String
    
    ''ProgressBar1.Visible = True
    
    If cmb_month.Text = "" Then
       MsgBox ("Select Month...")
       Exit Sub
    End If
    If cmb_year.Text = "" Then
       MsgBox ("Select Year...")
       Exit Sub
    End If
    find_dates
    
''    ProgressBar1.Value = ProgressBar1.Min
    
    Dim tablename, tablename2 As String
    tablename = "devicelogs_" + Trim(Str(cmb_month.ItemData(cmb_month.ListIndex))) + "_" + Trim(Str(cmb_year.Text))
    
    tablename2 = "devicelogs_" + Trim(Str(Month(end_date.Value))) + "_" + Trim(Str(Year(end_date.Value)))
    
    Dim dsnmdb As String
    Dim mdbrs As New ADODB.Recordset
    
    Set paydb = New ADODB.Connection
    Set payrs = New ADODB.Recordset
    
    paydb.Open pay

  ''  dsnmdb = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=\\192.168.11.31\eSSL\eTimeTrackLite\eTimeTrackLite1.mdb"
    dsnmdb = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=\\192.168.11.23\d\ESSL\etimetracklite\eTimeTrackLite1.mdb"
      
''    dsnmdb = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=\\192.168.11.9\ESSL\etimetracklite\eTimeTrackLite1.mdb"


''      dsnmdb = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=e:\eTimeTrackLite1.mdb"
''    dsnmdb = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=c:\eTimeTrackLite1.mdb"

    
   
    
    Dim woday As String


    pst_qry = "delete from bio_attendlogs where a_year = " & Val(cmb_year.Text) & "  and a_month = " & cmb_month.ItemData(cmb_month.ListIndex)
    paydb.Execute pst_qry


''    pst_qry = "delete from bio_devicelogs where ad_logdate between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and ad_auto = 'A'"
    pst_qry = "delete from bio_devicelogs where ad_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and ad_auto = 'A'"
    paydb.Execute pst_qry


    pst_qry = "delete from bio_device_shiftlogs where ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "'"
    paydb.Execute pst_qry


'''---select MSACESS MDB FILE

''    mdb_qry = "Select * from devicelogs as a, employees as b where a.deviceid <> 1 and a.userid = b.EmployeeCode and userid = '8007' and logdate between #" & Format(st_date, "MM/dd/yyyy") & "# and  #" & Format(end_date, "MM/dd/yyyy") & "# order by devicelogid"

''    mdb_qry = "Select * from devicelogs as a, employees as b where a.deviceid <> 1 and a.userid = b.EmployeeCode  and logdate between #" & Format(st_date, "MM/dd/yyyy") & "# and  #" & Format(end_date, "MM/dd/yyyy") & "# order by devicelogid "

    
    mdb_qry = "Select a.deviceid as deviceid , *  from devicelogs as a, employees as b where  a.deviceid <> 1 and a.userid = b.EmployeeCode  and logdate between #" & Format(st_date, "MM/dd/yyyy") & "# and  #" & Format(end_date + 1, "MM/dd/yyyy") & "# order by devicelogid "
    
    Dim logtype As String
    mdbrs.Open mdb_qry, dsnmdb, adOpenDynamic, adLockOptimistic
    While Not mdbrs.EOF
'         ProgressBar1.Value = ProgressBar1.Value + 1
          logtype = "A"
         If mdbrs!Deviceid = 1 Then
            logtype = "M"
         End If
         pst_qry = "insert into  bio_devicelogs (ad_fpcode,ad_empid,ad_logslno,ad_date,ad_logdate,ad_auto,ad_punch) values (  " & mdbrs!userid & ", " & mdbrs!employeeid & "," & mdbrs!devicelogid & ", '" & Format(mdbrs!logdate, "MM/dd/yyyy") & "' , '" & Format(mdbrs!logdate, "MM/dd/yyyy HH:MM:SS") & "' , '" & logtype & "','" & mdbrs!Direction & "' )"
         
         paydb.Execute pst_qry
         mdbrs.MoveNext
    Wend
    mdbrs.Close

'''

''    mdb_qry = "Select a.deviceid as deviceid , * from devicelogs_7_2017 as a, employees as b where  a.deviceid <> 1 and a.userid = b.EmployeeCode  and logdate between #" & Format(st_date, "MM/dd/yyyy") & "# and  #" & Format(end_date + 1, "MM/dd/yyyy") & "# order by devicelogid "
    
    mdb_qry = "Select a.deviceid as deviceid , * from " & tablename & " as a, employees as b where  a.deviceid <> 1 and a.userid = b.EmployeeCode  and logdate between #" & Format(st_date, "MM/dd/yyyy") & "# and  #" & Format(end_date + 1, "MM/dd/yyyy") & "# order by devicelogid "
    
    mdbrs.Open mdb_qry, dsnmdb, adOpenDynamic, adLockOptimistic
    While Not mdbrs.EOF
'         ProgressBar1.Value = ProgressBar1.Value + 1
          logtype = "A"
         If mdbrs!Deviceid = 1 Then
            logtype = "M"
         End If
         pst_qry = "insert into  bio_devicelogs (ad_fpcode,ad_empid,ad_logslno,ad_date,ad_logdate,ad_auto) values (  " & mdbrs!userid & ", " & mdbrs!employeeid & "," & mdbrs!devicelogid & ", '" & Format(mdbrs!logdate, "MM/dd/yyyy") & "' , '" & Format(mdbrs!logdate, "MM/dd/yyyy HH:MM:SS") & "' , '" & logtype & "' )"
         pst_qry = "insert into  bio_devicelogs (ad_fpcode,ad_empid,ad_logslno,ad_date,ad_logdate,ad_auto,ad_punch) values (  " & mdbrs!userid & ", " & mdbrs!employeeid & "," & mdbrs!devicelogid & ", '" & Format(mdbrs!logdate, "MM/dd/yyyy") & "' , '" & Format(mdbrs!logdate, "MM/dd/yyyy HH:MM:SS") & "' , '" & logtype & "','" & mdbrs!Direction & "' )"
         
         paydb.Execute pst_qry
         mdbrs.MoveNext
    Wend
    mdbrs.Close
    
    If Format(Now, "yyyy/MM/dd") >= Format(end_date.Value, "yyyy/MM/dd") Then

    If tablename <> tablename2 Then
        mdb_qry = "Select a.deviceid as deviceid , * from " & tablename2 & " as a, employees as b where  a.deviceid <> 1 and a.userid = b.EmployeeCode  and logdate between #" & Format(st_date, "MM/dd/yyyy") & "# and  #" & Format(end_date + 1, "MM/dd/yyyy") & "# order by devicelogid "
        
        mdbrs.Open mdb_qry, dsnmdb, adOpenDynamic, adLockOptimistic
        While Not mdbrs.EOF
    '         ProgressBar1.Value = ProgressBar1.Value + 1
              logtype = "A"
             If mdbrs!Deviceid = 1 Then
                logtype = "M"
             End If
             pst_qry = "insert into  bio_devicelogs (ad_fpcode,ad_empid,ad_logslno,ad_date,ad_logdate,ad_auto) values (  " & mdbrs!userid & ", " & mdbrs!employeeid & "," & mdbrs!devicelogid & ", '" & Format(mdbrs!logdate, "MM/dd/yyyy") & "' , '" & Format(mdbrs!logdate, "MM/dd/yyyy HH:MM:SS") & "' , '" & logtype & "' )"
             pst_qry = "insert into  bio_devicelogs (ad_fpcode,ad_empid,ad_logslno,ad_date,ad_logdate,ad_auto,ad_punch) values (  " & mdbrs!userid & ", " & mdbrs!employeeid & "," & mdbrs!devicelogid & ", '" & Format(mdbrs!logdate, "MM/dd/yyyy") & "' , '" & Format(mdbrs!logdate, "MM/dd/yyyy HH:MM:SS") & "' , '" & logtype & "','" & mdbrs!Direction & "' )"

             paydb.Execute pst_qry
             mdbrs.MoveNext
        Wend
        mdbrs.Close
    End If
    End If

''

    pst_qry = "insert into bio_attendlogs (a_bioid,a_fpcode,a_month,a_year) values (  '1027','1027', " & cmb_month.ItemData(cmb_month.ListIndex) & " , " & Val(cmb_year.Text) & " )"
    paydb.Execute pst_qry



    Dim idate, sft_from_date, sft_end_date As Date

   '' If rec_found = 0 Then
        mdb_qry = "Select * from employees where employeecode <> '0' and employeeid = 2819 and Status = 'Working' "
        mdb_qry = "Select * from employees where employeecode <> '0' and Status = 'Working' "
        mdbrs.Open mdb_qry, dsnmdb, adOpenDynamic, adLockOptimistic
        While Not mdbrs.EOF
             If Val(mdbrs!employeecode) > 0 Then
                pst_qry = "insert into bio_attendlogs (a_bioid,a_fpcode,a_month,a_year) values (  " & mdbrs!employeeid & "," & mdbrs!employeecode & ", " & cmb_month.ItemData(cmb_month.ListIndex) & " , " & Val(cmb_year.Text) & " )"
                paydb.Execute pst_qry
                
                
                For idate = st_date To end_date
                   pst_qry = "insert into  bio_device_shiftlogs (ds_fpcode,ds_empid,ds_date,ds_month,ds_year) values (  " & mdbrs!employeecode & ", " & mdbrs!employeeid & ",  '" & Format(idate, "MM/dd/yyyy") & "', " & cmb_month.ItemData(cmb_month.ListIndex) & " , " & Val(cmb_year.Text) & "   )"
                   paydb.Execute pst_qry
                Next
             End If
             mdbrs.MoveNext
        Wend
        mdbrs.Close
    
''for updating II
''Modified on 24/11/2015
''       dsnmdb = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=\\192.168.11.9\eSSL\eTimeTrackLite1.mdb"
''       mdb_qry = "Select * from employees where employeecode <> '0' and Status = 'Working' and team like '%CASUAL%' "
''        mdbrs.Open mdb_qry, dsnmdb, adOpenDynamic, adLockOptimistic
''        While Not mdbrs.EOF
''             pst_qry = "insert into bio_attendlogs (a_bioid,a_fpcode,a_month,a_year) values (  " & mdbrs!employeeid & "," & mdbrs!employeecode & ", " & cmb_month.ItemData(cmb_month.ListIndex) & " , " & Val(cmb_year.Text) & " )"
''             paydb.Execute pst_qry
''             For idate = st_date To end_date
''                pst_qry = "insert into  bio_device_shiftlogs (ds_fpcode,ds_empid,ds_date,ds_month,ds_year) values (  " & mdbrs!employeecode & ", " & mdbrs!employeeid & ",  '" & Format(idate, "MM/dd/yyyy") & "', " & cmb_month.ItemData(cmb_month.ListIndex) & " , " & Val(cmb_year.Text) & "   )"
''                paydb.Execute pst_qry
''
''             Next
''             mdbrs.MoveNext
''        Wend
''        mdbrs.Close

    pst_qry = "select *  from mas_caemp where ca_biometric = 'Y'"
    
    pst_qry = "select *  from mas_caemp where ca_biometric = 'Y' and ca_fpcode not in (select a_fpcode from bio_attendlogs where a_month = " & cmb_month.ItemData(cmb_month.ListIndex) & "   and a_year = " & Val(cmb_year.Text) & " )"
    payrs.Open pst_qry, paydb, adOpenDynamic, adLockOptimistic
    While Not payrs.EOF
          pst_qry = "insert into bio_attendlogs (a_bioid,a_fpcode,a_month,a_year) values ( 00 ," & payrs!ca_fpcode & ", " & cmb_month.ItemData(cmb_month.ListIndex) & " , " & Val(cmb_year.Text) & " )"
          paydb.Execute pst_qry
          For idate = st_date To end_date
              pst_qry = "insert into  bio_device_shiftlogs (ds_fpcode,ds_empid,ds_date,ds_month,ds_year) values (  " & payrs!ca_fpcode & ", 0,  '" & Format(idate, "MM/dd/yyyy") & "', " & cmb_month.ItemData(cmb_month.ListIndex) & " , " & Val(cmb_year.Text) & "   )"
              paydb.Execute pst_qry
          Next
          payrs.MoveNext
    Wend
    payrs.Close

    
    
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

Private Sub cmd_update_all_Click()
    all_chk = 0
    upload_biometric_data
    update_leave_od_data
End Sub

Private Sub cmd_update_leave_od_Click()
         
    pst_ans = MsgBox("Are u sure want to Update Leave / OD Particulars ...  ", vbYesNo)
    If pst_ans = vbNo Then Exit Sub
   
    For idate = st_date To end_date
         pst_qry = "insert into  bio_device_shiftlogs (ds_fpcode,ds_empid,ds_date,ds_month,ds_year) values (  '1027', '1027',  '" & Format(idate, "MM/dd/yyyy") & "', " & cmb_month.ItemData(cmb_month.ListIndex) & " , " & Val(cmb_year.Text) & "   )"
         paydb.Execute pst_qry
    Next
   
''    sql = "select * from payroll_lock where pay_finyear = " & finyear & " and pay_month = " & cmb_month.ItemData(cmb_month.ListIndex) & " and pay_year = " & Val(cmb_year.Text) & ""
''    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
''    If Not payrs.EOF Then
''      MsgBox ("Attendance details are locked.. You can't Update Again...")
''      payrs.Close
''      Exit Sub
''    End If
    all_chk = 1

    pst_qry = "update bio_devicelogs set ad_upd ='N' where ad_date between '" & Format(st_date - 1, "MM/dd/yyyy") & "' and  '" & Format(end_date + 1, "MM/dd/yyyy") & "'"
    paydb.Execute pst_qry

   all_chk = 1
   update_leave_od_data
   cmd_exit.SetFocus
   
End Sub

Private Sub cmd_upload_biometric_Click()
   pst_ans = MsgBox("Are u sure want to Update Attendance Logs from Bio-metric machines ...  ", vbYesNo)
   If pst_ans = vbNo Then Exit Sub
   
   
   sql = "select * from payroll_lock where pay_finyear = " & finyear & " and pay_month = " & cmb_month.ItemData(cmb_month.ListIndex) & " and pay_year = " & Val(cmb_year.Text) & ""
   payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
   If Not payrs.EOF Then
      MsgBox ("Attendance details are locked.. You can't Update Again...")
      Exit Sub
   End If
   all_chk = 1
   
   upload_biometric_data
       
''    For idate = st_date To end_date
''         pst_qry = "insert into  bio_device_shiftlogs (ds_fpcode,ds_empid,ds_date,ds_month,ds_year) values (  '1027', '1027',  '" & Format(idate, "MM/dd/yyyy") & "', " & cmb_month.ItemData(cmb_month.ListIndex) & " , " & Val(cmb_year.Text) & "   )"
''         paydb.Execute pst_qry
''    Next

End Sub

Private Sub cmd_wo_upd_Click()
        If cmb_month.Text = "" Then
           MsgBox ("Select Month...")
           Exit Sub
        End If
        If cmb_year.Text = "" Then
           MsgBox ("Select Year...")
           Exit Sub
        End If
        find_dates


    
    Dim dsnmdb As String
    Dim mdbrs As New ADODB.Recordset
    
    Set paydb = New ADODB.Connection
    Set payrs = New ADODB.Recordset
    
    paydb.Open pay
    
    
    
    
    
    

    dsnmdb = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=\\192.168.11.31\eSSL\eTimeTrackLite\eTimeTrackLite1.mdb"
    dsnmdb = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=\\192.168.11.23\d\ESSL\etimetracklite\eTimeTrackLite1.mdb"

'' dsnmdb = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=\\192.168.11.9\dpm\eTimeTrackLite1.mdb"
    
''    dsnmdb = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=c:\eTimeTrackLite1.mdb"

'''---select MSACESS MDB FILE

''    mdb_qry = "Select * from devicelogs as a, employees as b where a.deviceid <> 1 and a.userid = b.EmployeeCode and userid = '8007' and logdate between #" & Format(st_date, "MM/dd/yyyy") & "# and  #" & Format(end_date, "MM/dd/yyyy") & "# order by devicelogid"






'' for updating WO DATE
    
    Dim leave As String
    Dim empid As String
    mdb_qry = "select * from employees as a , categories as b where a.categoryId = b.categoryId  and status = 'Working'"
    mdbrs.Open mdb_qry, dsnmdb, adOpenDynamic, adLockOptimistic
    While Not mdbrs.EOF
           leave = mdbrs!categorysname
           id = mdbrs!employeecode
           If id > 0 Then
              pst_qry = "update emp_mas set emp_holiday = '" & leave & "' where emp_fpcode = " & id & " and emp_status = 'A'"
              paydb.Execute pst_qry
           End If
          mdbrs.MoveNext
    Wend
    mdbrs.Close
    
    
    MsgBox ("updated..")

    paydb.Close
    
End Sub

Private Sub exit_Click()
    Unload Me
End Sub


Private Sub Form_Load()
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
    
    
''    ProgressBar1.Visible = False
    cmb_month.ListIndex = Month(Date) - 1
End Sub



''''Private Sub save_Click()
''''On Error GoTo err_handler
''''
''''    Dim id, fcode As Integer
''''    Dim dlogdate As Date
''''
''''    Dim dev_log(100) As Long
''''
''''    Dim log_details As String
''''
''''
''''    Dim stime, etime As Date
''''
''''    stime = TimeValue(Now)
''''
''''
''''    Dim sft, sft_bt, sft_et, sft_begin_dur, sft_end_dur As String
''''
''''    ''ProgressBar1.Visible = True
''''
''''    If Opt2.Value = True Then
''''        If cmb_month.Text = "" Then
''''           MsgBox ("Select Month...")
''''           Exit Sub
''''        End If
''''        If cmb_year.Text = "" Then
''''           MsgBox ("Select Year...")
''''           Exit Sub
''''        End If
''''        find_dates
''''    Else
''''        end_date = dt_ason.Value
''''        st_date = dt_ason.Value
''''
''''    End If
''''
''''    ProgressBar1.Value = ProgressBar1.Min
''''
''''    Dim dsnmdb As String
''''    Dim mdbrs As New ADODB.Recordset
''''
''''    Set paydb = New ADODB.Connection
''''    Set payrs = New ADODB.Recordset
''''
''''    paydb.Open pay
''''
''''    dsnmdb = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=\\192.168.11.31\eSSL\eTimeTrackLite\eTimeTrackLite1.mdb"
''''
''''''    dsnmdb = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=c:\eTimeTrackLite1.mdb"
''''
''''
''''''    paydb.BeginTrans
'''''start
''''''    Dim rec_found As Integer
''''''    pst_qry = "select  count(*) as totrecs from TEMPbio_device_shiftlogs where ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "'"
''''''    payrs.Open pst_qry, paydb, 1, 2
''''''    If Not payrs.EOF = True Then
''''''       rec_found = payrs("totrecs")
''''''    End If
''''''    payrs.Close
''''
''''
''''    Dim woday As String
''''
''''
''''    pst_qry = "delete from bio_attendlogs where a_year = " & Val(cmb_year.Text) & "  and a_month = " & cmb_month.ItemData(cmb_month.ListIndex)
''''    paydb.Execute pst_qry
''''
''''
''''    pst_qry = "delete from bio_devicelogs where ad_logdate between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "'"
''''    paydb.Execute pst_qry
''''
''''
''''    pst_qry = "delete from TEMPbio_device_shiftlogs where ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "'"
''''    paydb.Execute pst_qry
''''
''''
'''''''---select MSACESS MDB FILE
''''
''''''    mdb_qry = "Select * from devicelogs as a, employees as b where a.deviceid <> 1 and a.userid = b.EmployeeCode and userid = '8007' and logdate between #" & Format(st_date, "MM/dd/yyyy") & "# and  #" & Format(end_date, "MM/dd/yyyy") & "# order by devicelogid"
''''
''''    mdb_qry = "Select * from devicelogs as a, employees as b where a.deviceid <> 1 and a.userid = b.EmployeeCode  and logdate between #" & Format(st_date, "MM/dd/yyyy") & "# and  #" & Format(end_date, "MM/dd/yyyy") & "# order by devicelogid "
''''
''''    mdb_qry = "Select * from devicelogs as a, employees as b where a.userid = b.EmployeeCode  and logdate between #" & Format(st_date, "MM/dd/yyyy") & "# and  #" & Format(end_date, "MM/dd/yyyy") & "# order by devicelogid "
''''
''''    Dim logtype As String
''''
''''    mdbrs.Open mdb_qry, dsnmdb, adOpenDynamic, adLockOptimistic
''''    While Not mdbrs.EOF
'''''         ProgressBar1.Value = ProgressBar1.Value + 1
''''          logtype = "A"
''''         If mdbrs!deviceid = 1 Then
''''            logtype = "M"
''''         End If
''''
''''         pst_qry = "insert into  bio_devicelogs (ad_fpcode,ad_empid,ad_logslno,ad_date,ad_logdate,ad_auto) values (  " & mdbrs!userid & ", " & mdbrs!employeeid & "," & mdbrs!devicelogid & ", '" & Format(mdbrs!logdate, "MM/dd/yyyy") & "' , '" & Format(mdbrs!logdate, "MM/dd/yyyy HH:MM:SS") & "' , '" & logtype & "' )"
''''         paydb.Execute pst_qry
''''         mdbrs.MoveNext
''''    Wend
''''    mdbrs.Close
''''
''''
''''    Dim idate, sft_from_date, sft_end_date As Date
''''
''''   '' If rec_found = 0 Then
''''        mdb_qry = "Select * from employees where employeecode <> '0' and employeeid = 2819 and Status = 'Working' "
''''        mdb_qry = "Select * from employees where employeecode <> '0' and Status = 'Working' "
''''        mdbrs.Open mdb_qry, dsnmdb, adOpenDynamic, adLockOptimistic
''''        While Not mdbrs.EOF
''''             pst_qry = "insert into bio_attendlogs (a_bioid,a_fpcode,a_month,a_year) values (  " & mdbrs!employeeid & "," & mdbrs!employeecode & ", " & cmb_month.ItemData(cmb_month.ListIndex) & " , " & Val(cmb_year.Text) & " )"
''''             paydb.Execute pst_qry
''''             For idate = st_date To end_date
''''                pst_qry = "insert into  bio_device_shiftlogs (ds_fpcode,ds_empid,ds_date) values (  " & mdbrs!employeecode & ", " & mdbrs!employeeid & ",  '" & Format(idate, "MM/dd/yyyy") & "'  )"
''''                paydb.Execute pst_qry
''''
''''             Next
''''             mdbrs.MoveNext
''''        Wend
''''        mdbrs.Close
''''
''''        Dim sftfound As Integer
''''        sql = "update bio_device_shiftlogs set ds_shift = 'GS',ds_shift_begintime = '08:00',ds_shift_endtime = '17:00',ds_begin_duration  = '60',ds_end_duration = '420' where ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' "
''''        paydb.Execute sql
''''
''''''    End If
''''
''''
''''
''''''Employee shift updation
''''''    mdb_qry = "Select * from employeeshift as a, shifts as b where a.shiftid = b.shiftid  and fromdate >=  #" & Format(st_date, "dd/MM/yyyy") & "# and todate <=  #" & Format(end_date, "dd/MM/yyyy") & "# "
''''''    mdbrs.Open mdb_qry, dsnmdb, adOpenDynamic, adLockOptimistic
''''''    While Not mdbrs.EOF
''''''          id = mdbrs!employeeid
''''''          sft = mdbrs!shiftsname
''''''          sft_bt = mdbrs!begintime
''''''          sft_et = mdbrs!endtime
''''''          sft_begin_dur = mdbrs!punchbeginduration
''''''          sft_end_dur = mdbrs!punchendduration
''''''          sft_from_date = mdbrs!fromdate
''''''          sft_end_date = mdbrs!todate
''''''          sql = "update bio_device_shiftlogs set ds_shift = '" & sft & "',ds_shift_begintime = '" & sft_bt & "',ds_shift_endtime = '" & sft_et & "',ds_begin_duration  = '" & sft_begin_dur & "',ds_end_duration = '" & sft_begin_dur & "' where ds_empid =  '" & id & "' and ds_date between '" & Format(sft_from_date, "MM/dd/yyyy") & "' and '" & Format(sft_end_date, "MM/dd/yyyy") & "' "
''''''          paydb.Execute sql
''''''          mdbrs.MoveNext
''''''    Wend
''''''    mdbrs.Close
''''
''''    pst_qry = "select *  from bio_shift_schedule where emps_date  between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "'"
''''    payrs.Open pst_qry, paydb, adOpenDynamic, adLockOptimistic
''''    While Not payrs.EOF
''''          id = payrs!emps_fpcode
''''          sft = payrs!emps_shift
''''          sft_from_date = payrs!emps_date
''''
''''''          sql = "update bio_device_shiftlogs set ds_shift = '" & sft & "',ds_shift_begintime = '" & sft_bt & "',ds_shift_endtime = '" & sft_et & "',ds_begin_duration  = '" & sft_begin_dur & "',ds_end_duration = '" & sft_begin_dur & "' where ds_empid =  '" & id & "' and ds_date between '" & Format(sft_from_date, "MM/dd/yyyy") & "' and '" & Format(sft_end_date, "MM/dd/yyyy") & "' "
''''          sql = "update bio_device_shiftlogs set ds_shift = '" & sft & "' where ds_fpcode =  '" & id & "' and ds_date =  '" & Format(sft_from_date, "MM/dd/yyyy") & "'"
''''
''''          paydb.Execute sql
''''          payrs.MoveNext
''''    Wend
''''    payrs.Close
''''
''''
''''
''''
''''
''''
''''''    Dim leave, leavetype As String
'''' ''for updating leave entries
''''
''''''    mdb_qry = "Select * from leaveentries as a, leavetypes as b where a.leavetypeid = b.leavetypeid  and fromdate >=  #" & Format(st_date, "MM/dd/yyyy") & "# and todate <=  #" & Format(end_date, "MM/dd/yyyy") & "# order by employeeid , fromdate "
''''''    mdbrs.Open mdb_qry, dsnmdb, adOpenDynamic, adLockOptimistic
''''''    While Not mdbrs.EOF
''''''           leave = ""
''''''           If mdbrs!leavestatus = "HalfDay" Then
''''''              leave = "½"
''''''           End If
''''''           sft_from_date = mdbrs!fromdate
''''''           sft_end_date = mdbrs!todate
''''''           leavetype = leave + mdbrs!leavetypesname
''''''           id = mdbrs!employeeid
''''''''           If id = 2642 Then
''''''''              MsgBox ("Wait")
''''''''           End If
''''''
''''''           sql = "update bio_device_shiftlogs set ds_status = '" & leavetype & "' where ds_empid =  '" & id & "' and ds_date between '" & Format(sft_from_date, "MM/dd/yyyy") & "' and '" & Format(sft_end_date, "MM/dd/yyyy") & "' "
''''''           paydb.Execute sql
''''''           mdbrs.MoveNext
''''''    Wend
''''''    mdbrs.Close
''''
''''''
''''''
''''''    pst_qry = "select * from bio_empleave where emp_leave_date between '" & Format(st_date, "MM/dd/yyyy") & "' and '" & Format(end_date, "MM/dd/yyyy") & "'"
''''''
''''''    payrs.Open pst_qry, paydb, adOpenDynamic, adLockOptimistic
''''''    While Not payrs.EOF
''''''          id = payrs!emp_fpcode
''''''          leavetype = payrs!emp_leave_type
''''''          leave = IIf(payrs!emp_leave_period = "F", "EL", "½EL")
''''''          sql = "update bio_device_shiftlogs set ds_status = '" & leave & "' where ds_fpcode =  '" & id & "' and ds_date =  '" & Format(payrs!emp_leave_date, "MM/dd/yyyy") & "'"
''''''          paydb.Execute sql
''''''          payrs.MoveNext
''''''    Wend
''''''    payrs.Close
''''''
''''
''''
'''' ''for updating OD entries
''''''    mdb_qry = "Select * from specialentries where  fromdate >=  #" & Format(st_date, "MM/dd/yyyy") & "# and todate <=  #" & Format(end_date, "MM/dd/yyyy") & "# order by employeeid , fromdate "
''''''    mdbrs.Open mdb_qry, dsnmdb, adOpenDynamic, adLockOptimistic
''''''    While Not mdbrs.EOF
''''''           sft_from_date = mdbrs!fromdate
''''''           sft_end_date = mdbrs!todate
''''''           id = mdbrs!employeeid
''''''           sql = "update bio_device_shiftlogs set ds_status = 'P(OD)' where ds_empid =  '" & id & "' and ds_date between '" & Format(sft_from_date, "MM/dd/yyyy") & "' and '" & Format(sft_end_date, "MM/dd/yyyy") & "' "
''''''           paydb.Execute sql
''''''           mdbrs.MoveNext
''''''    Wend
''''''    mdbrs.Close
''''
''''
''''
''''''for updating general shift , A Shift , B Shift
''''    pst_qry = "Update bio_device_shiftlogs set ds_shift_in = intime, ds_shift_out = outtime  from bio_device_shiftlogs a , (select ad_fpcode,ad_empid,ad_date, min(ad_logdate) as intime , max(ad_logdate) as outtime    from bio_devicelogs  group by ad_fpcode,ad_empid,ad_date ) b Where ds_shift in ('GS','A SHIFT','B SHIFT','6.00 to 6.00 (Day)') and ds_fpcode = ad_fpcode And ds_empid = ad_empid And ds_date = ad_date and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "'"
''''    paydb.Execute pst_qry
''''
''''    Set paydb2 = New ADODB.Connection
''''    Set payrs2 = New ADODB.Recordset
''''    paydb2.Open pay
''''    Dim i, j As Integer
''''    Dim dt_log As Integer
''''
''''''for updating B Shift - 02.00 PM - 10.00 PM
''''    sql = "select * from bio_device_shiftlogs where ds_shift = 'B SHIFT' and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "'"
''''    payrs.Open sql, paydb, 1, 2
''''    While Not payrs.EOF
''''         dt_log = 1
''''          i = 1
''''          id = payrs!ds_empid
''''          sft_from_date = payrs!ds_date
''''          pst_qry = "select ad_date,ad_logslno,min(ad_logdate) as gtime from bio_devicelogs where ad_upd  = 'N' and ad_fpcode =  " & id & "  and ad_date >= '" & Format(sft_from_date, "MM/dd/yyyy") & "' and ad_date <= '" & Format(sft_from_date + 1, "MM/dd/yyyy") & "' and ad_upd = 'N' group by ad_date,ad_logslno"
''''          payrs2.Open pst_qry, paydb2, 1, 2
''''          While Not payrs2.EOF
''''              If i = 1 Then
''''                 payrs("ds_shift_in") = payrs2!gtime
''''                 dev_log(dt_log) = payrs2!ad_logslno
''''                 dt_log = dt_log + 1
''''
''''              ElseIf i = 2 Then
''''                 payrs("ds_shift_out") = payrs2!gtime
''''                 dev_log(dt_log) = payrs2!ad_logslno
''''                 dt_log = dt_log + 1
''''
''''              End If
''''              i = i + 1
''''              payrs2.MoveNext
''''          Wend
''''          payrs2.Close
''''          payrs.Update
''''
''''          log_details = "(0"
''''          For j = 1 To dt_log - 1
''''              log_details = log_details + "," + Str(dev_log(j))
''''          Next
''''          log_details = log_details + ")"
''''          pst_qry = "update bio_devicelogs set ad_upd  = 'Y' where  ad_fpcode =  " & id & "  and ad_logslno in " & log_details
''''          paydb2.Execute pst_qry
''''
''''          payrs.MoveNext
''''    Wend
''''    payrs.Close
''''
''''
''''
''''
''''
''''
''''''for updating B&C Shift - 6.00 PM - 6.00 AM and   C SHIFT - 2.00 PM - 6.00 AM
''''    sql = "select * from bio_device_shiftlogs where ds_shift = '6.00 to 6.00 (Night)' and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "'"
''''
''''    sql = "select * from bio_device_shiftlogs where ds_shift in ('06.00PM-06.00AM','6.00 to 6.00 (Night)','C SHIFT') and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "'"
''''    payrs.Open sql, paydb, 1, 2
''''    While Not payrs.EOF
''''          dt_log = 1
''''          i = 1
''''          id = payrs!ds_empid
''''          sft_from_date = payrs!ds_date
''''
''''''          pst_qry = "select ad_date,ad_logslno,min(ad_logdate) as gtime from bio_devicelogs where ad_fpcode =  " & id & "  and ad_date >= '" & Format(sft_from_date, "MM/dd/yyyy") & "' and ad_date <= '" & Format(sft_from_date + 1, "MM/dd/yyyy") & "' group by ad_date,ad_logslno"
''''
''''          pst_qry = "select ad_date,ad_logslno,gtime from " _
''''              & " (select ad_date,ad_logslno,min(ad_logdate) as gtime from bio_devicelogs where ad_fpcode =  " & id & "  and ad_date = '" & Format(sft_from_date, "MM/dd/yyyy") & "' and datepart(hh,ad_logdate) > 15  and ad_upd = 'N'  group by ad_date,ad_logslno " _
''''              & " Union All " _
''''              & " select ad_date,ad_logslno,min(ad_logdate) as gtime from bio_devicelogs where ad_fpcode =  " & id & "  and ad_date = '" & Format(sft_from_date + 1, "MM/dd/yyyy") & "' and datepart(hh,ad_logdate) <10  and ad_upd = 'N' group by ad_date,ad_logslno ) a group by  ad_date,ad_logslno,gtime "
''''
''''''          pst_qry = "select ad_date,ad_logslno,gtime,ad_upd from " _
''''''              & " (select ad_date,ad_logslno,min(ad_logdate) as gtime,ad_upd from bio_devicelogs where ad_fpcode =  " & id & "  and ad_date = '" & Format(sft_from_date, "MM/dd/yyyy") & "' and datepart(hh,ad_logdate) > 15  group by ad_date,ad_logslno,ad_upd " _
''''''              & " Union All " _
''''''              & " select ad_date,ad_logslno,min(ad_logdate) as gtime,ad_upd from bio_devicelogs where ad_fpcode =  " & id & "  and ad_date = '" & Format(sft_from_date + 1, "MM/dd/yyyy") & "' and datepart(hh,ad_logdate) <7  group by ad_date,ad_logslno,ad_upd ) a group by  ad_date,ad_logslno,gtime,ad_upd "
''''
''''
''''          payrs2.Open pst_qry, paydb2, 1, 2
''''          While Not payrs2.EOF
''''              If i = 1 Then
''''                 payrs("ds_shift_in") = payrs2!gtime
''''                 dev_log(dt_log) = payrs2!ad_logslno
''''                 dt_log = dt_log + 1
''''
''''              ElseIf i = 2 Then
''''                 payrs("ds_shift_out") = payrs2!gtime
''''                 dev_log(dt_log) = payrs2!ad_logslno
''''                 dt_log = dt_log + 1
''''              End If
''''              i = i + 1
''''
''''              payrs2.MoveNext
''''          Wend
''''          payrs2.Close
''''          payrs.Update
''''
''''          log_details = "(0"
''''          For j = 1 To dt_log - 1
''''              log_details = log_details + "," + Str(dev_log(j))
''''          Next
''''          log_details = log_details + ")"
''''          pst_qry = "update bio_devicelogs set ad_upd  = 'Y' where  ad_fpcode =  " & id & "  and ad_logslno in " & log_details
''''          paydb2.Execute pst_qry
''''          payrs.MoveNext
''''    Wend
''''    payrs.Close
''''
''''''for updating B+C Shift - 2.00 PM - NEXT DAY 6.00 AM
''''''    dt_log = 1
''''
''''    sql = "select * from bio_device_shiftlogs where ds_shift = 'B+C' and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "'"
''''    payrs.Open sql, paydb, 1, 2
''''    While Not payrs.EOF
''''          dt_log = 1
''''          i = 1
''''          id = payrs!ds_empid
''''          sft_from_date = payrs!ds_date
''''
''''''          pst_qry = "select ad_date,ad_logslno,min(ad_logdate) as gtime from bio_devicelogs where ad_fpcode =  " & id & "  and ad_date >= '" & Format(sft_from_date, "MM/dd/yyyy") & "' and ad_date <= '" & Format(sft_from_date + 1, "MM/dd/yyyy") & "' group by ad_date,ad_logslno"
''''
''''          pst_qry = "select ad_date,ad_logslno,gtime from " _
''''              & " (select ad_date,ad_logslno,min(ad_logdate) as gtime from bio_devicelogs where ad_fpcode =  " & id & "  and ad_date = '" & Format(sft_from_date, "MM/dd/yyyy") & "' and datepart(hh,ad_logdate) > 13  and ad_upd = 'N' group by ad_date,ad_logslno " _
''''              & " Union All " _
''''              & " select ad_date,ad_logslno,min(ad_logdate) as gtime from bio_devicelogs where ad_fpcode =  " & id & "  and ad_date = '" & Format(sft_from_date + 1, "MM/dd/yyyy") & "' and datepart(hh,ad_logdate) <7  and ad_upd = 'N' group by ad_date,ad_logslno ) a group by  ad_date,ad_logslno,gtime "
''''
''''
''''          payrs2.Open pst_qry, paydb2, 1, 2
''''          While Not payrs2.EOF
''''              If i = 1 Then
''''                 payrs("ds_shift_in") = payrs2!gtime
''''                 dev_log(dt_log) = payrs2!ad_logslno
''''                 dt_log = dt_log + 1
''''              ElseIf i = 2 Then
''''                 payrs("ds_shift_out") = payrs2!gtime
''''                 dev_log(dt_log) = payrs2!ad_logslno
''''                 dt_log = dt_log + 1
''''              End If
''''              i = i + 1
''''  ''            payrs2("ad_upd") = "Y"
''''  ''            payrs2.Update
''''              payrs2.MoveNext
''''          Wend
''''          payrs2.Close
''''          payrs.Update
''''          log_details = "(0"
''''          For j = 1 To dt_log - 1
''''             log_details = log_details + "," + Str(dev_log(j))
''''          Next
''''          log_details = log_details + ")"
''''          pst_qry = "update bio_devicelogs set ad_upd  = 'Y' where  ad_fpcode =  " & id & "  and ad_logslno in " & log_details
''''          paydb2.Execute pst_qry
''''
''''          payrs.MoveNext
''''    Wend
''''    payrs.Close
'''''end
''''
''''''for updating A+B+C Shift - 6.00 AM - NEXT DAY 6.00 AM
''''''    dt_log = 1
''''
''''    sql = "select * from bio_device_shiftlogs where ds_shift = 'A+B+C' and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "'"
''''    payrs.Open sql, paydb, 1, 2
''''    While Not payrs.EOF
''''          dt_log = 1
''''          i = 1
''''          id = payrs!ds_empid
''''          sft_from_date = payrs!ds_date
''''
''''''          pst_qry = "select ad_date,ad_logslno,min(ad_logdate) as gtime from bio_devicelogs where ad_fpcode =  " & id & "  and ad_date >= '" & Format(sft_from_date, "MM/dd/yyyy") & "' and ad_date <= '" & Format(sft_from_date + 1, "MM/dd/yyyy") & "' group by ad_date,ad_logslno"
''''
''''          pst_qry = "select ad_date,ad_logslno,gtime from " _
''''              & " (select ad_date,ad_logslno,min(ad_logdate) as gtime from bio_devicelogs where ad_fpcode =  " & id & "  and ad_date = '" & Format(sft_from_date, "MM/dd/yyyy") & "' and datepart(hh,ad_logdate) > 5  and ad_upd = 'N' group by ad_date,ad_logslno " _
''''              & " Union All " _
''''              & " select ad_date,ad_logslno,min(ad_logdate) as gtime from bio_devicelogs where ad_fpcode =  " & id & "  and ad_date = '" & Format(sft_from_date + 1, "MM/dd/yyyy") & "' and datepart(hh,ad_logdate) <7  and ad_upd = 'N' group by ad_date,ad_logslno ) a group by  ad_date,ad_logslno,gtime "
''''
''''
''''          payrs2.Open pst_qry, paydb2, 1, 2
''''          While Not payrs2.EOF
''''              If i = 1 Then
''''                 payrs("ds_shift_in") = payrs2!gtime
''''                 dev_log(dt_log) = payrs2!ad_logslno
''''                 dt_log = dt_log + 1
''''              ElseIf i = 2 Then
''''                 payrs("ds_shift_out") = payrs2!gtime
''''                 dev_log(dt_log) = payrs2!ad_logslno
''''                 dt_log = dt_log + 1
''''              End If
''''              i = i + 1
''''  ''            payrs2("ad_upd") = "Y"
''''  ''            payrs2.Update
''''              payrs2.MoveNext
''''          Wend
''''          payrs2.Close
''''          payrs.Update
''''          log_details = "(0"
''''          For j = 1 To dt_log - 1
''''             log_details = log_details + "," + Str(dev_log(j))
''''          Next
''''          log_details = log_details + ")"
''''          pst_qry = "update bio_devicelogs set ad_upd  = 'Y' where  ad_fpcode =  " & id & "  and ad_logslno in " & log_details
''''          paydb2.Execute pst_qry
''''
''''          payrs.MoveNext
''''    Wend
''''    payrs.Close
'''''end
''''
''''
''''
''''''Weekoff updation for sundays
''''''         sql = "update bio_device_shiftlogs set ds_shift = 'WO',ds_shift_begintime = '',ds_shift_endtime = '',ds_begin_duration  = '',ds_end_duration = '' where ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and datepart(dw,ds_date) =  1"
''''  ''       paydb.Execute sql
''''
''''
''''''    Dim woday As String
''''    pst_qry = "select * from emp_mas where emp_company in (1,2,3,5) and emp_status = 'A' "
''''    payrs.Open pst_qry, paydb, 1, 2
''''    While Not payrs.EOF
''''         If payrs("emp_holiday") = "SUNDAY" Then
''''                woday = 1
''''         ElseIf payrs("emp_holiday") = "MONDAY" Then
''''                woday = 2
''''         ElseIf payrs("emp_holiday") = "TUESDAY" Then
''''                woday = 3
''''         ElseIf payrs("emp_holiday") = "WEDNESDAY" Then
''''                woday = 4
''''         ElseIf payrs("emp_holiday") = "THURSDAY" Then
''''                woday = 5
''''         ElseIf payrs("emp_holiday") = "FRIDAY" Then
''''                woday = 6
''''         ElseIf payrs("emp_holiday") = "SATURDAY" Then
''''                woday = 7
''''         Else
''''                woday = 0
''''         End If
''''''         If payrs("emp_fpcode") = 1093 Then
''''''            MsgBox (payrs("emp_fpcode"))
''''''         End If
''''         sql = "update bio_device_shiftlogs set ds_status = 'WO', ds_shift = 'WO',ds_shift_begintime = '',ds_shift_endtime = '',ds_begin_duration  = '',ds_end_duration = '' where ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and ds_fpcode = " & payrs("emp_fpcode") & " and datepart(dw,ds_date) =  " & woday
''''         paydb.Execute sql
''''         payrs.MoveNext
''''    Wend
''''    payrs.Close
''''''    sql = "update bio_device_shiftlogs set ds_status = 'WO' WHERE   ds_shift = 'WO' and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "'"
''''''    paydb.Execute sql
''''
''''''for updating Declared holiday
''''''start
''''    pst_qry = "Select * from emp_dec_holiday where emp_dec_holiday between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "'"
''''    payrs.Open pst_qry, paydb, 1, 2
''''    While Not payrs.EOF
''''          sql = "update bio_device_shiftlogs set ds_shift = 'H',ds_shift_begintime = '00:00',ds_shift_endtime = '00:00',ds_begin_duration  = '',ds_end_duration = '' where ds_date = '" & Format(payrs!emp_dec_holiday, "MM/dd/yyyy") & "'"
''''          paydb.Execute sql
''''          payrs.MoveNext
''''    Wend
''''    payrs.Close
''''''end
''''
''''    sql = "update bio_device_shiftlogs set ds_sft_hrs = (case when DATEpart(minute,ds_shift_out) > 40 then datepart(hour,ds_shift_out)+1+datediff(day,ds_shift_in,ds_shift_out)*24  else datepart(hour,ds_shift_out)+datediff(day,ds_shift_in,ds_shift_out)*24 end - case when DATEpart(minute,ds_shift_in) > 30 then datepart(hour,ds_shift_in)+1 else datepart(hour,ds_shift_in) end) where ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "'"
''''    paydb.Execute sql
''''
''''
''''
''''    Dim ftime As String
'''' ''for updating Permission entries
''''    pst_qry = "select * from bio_emp_permissions where empp_date between '" & Format(st_date, "MM/dd/yyyy") & "' and '" & Format(end_date, "MM/dd/yyyy") & "'"
''''    payrs.Open pst_qry, paydb, adOpenDynamic, adLockOptimistic
''''    While Not payrs.EOF
''''          id = payrs!empp_fpcode
''''          idate = payrs!empp_date
''''
''''          ftime = Val(payrs!empp_fromtime)
''''          etime = Val(payrs!empp_endtime)
''''
''''
''''''          sql = "update bio_device_shiftlogs set ds_status = '" & leave & "' where ds_fpcode =  '" & id & "' and ds_date =  '" & Format(payrs!emp_leave_date, "MM/dd/yyyy") & "'"
''''''          paydb.Execute sql
''''          payrs.MoveNext
''''    Wend
''''    payrs.Close
''''
''''
''''    sql = "update bio_device_shiftlogs set ds_status = 'A' WHERE ds_sft_hrs = 0 and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(Now - 1, "MM/dd/yyyy") & "' and ds_shift <> 'WO' "
''''    paydb.Execute sql
''''
''''    sql = "update bio_device_shiftlogs set ds_status = 'H' WHERE   ds_shift = 'H' and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "'"
''''    paydb.Execute sql
''''
''''
''''
''''    sql = "update bio_device_shiftlogs set ds_status = '½HP' WHERE ds_sft_hrs > 4 and ds_sft_hrs < 7 and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and  ds_shift = 'H'"
''''    paydb.Execute sql
''''
''''    sql = "update bio_device_shiftlogs set ds_status = 'HP' WHERE ds_sft_hrs > 7.30 and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and  ds_shift = 'H'"
''''    paydb.Execute sql
''''
''''
''''    sql = "update bio_device_shiftlogs set ds_status = 'P' WHERE ds_sft_hrs > 6 and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "'  and  ds_shift <> 'H'"
''''    paydb.Execute sql
''''
''''
''''    sql = "update bio_device_shiftlogs set ds_status = 'WOP' WHERE ds_sft_hrs > 7 and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and  ds_shift = 'WO' and ds_status = 'P'"
''''    paydb.Execute sql
''''
''''
''''    Dim leave, leavetype As String
'''' ''for updating leave entries
''''    pst_qry = "select * from bio_empleave where emp_leave_date between '" & Format(st_date, "MM/dd/yyyy") & "' and '" & Format(end_date, "MM/dd/yyyy") & "'"
''''    payrs.Open pst_qry, paydb, adOpenDynamic, adLockOptimistic
''''    While Not payrs.EOF
''''          id = payrs!emp_fpcode
''''          leavetype = payrs!emp_leave_type
''''          If payrs!emp_leave_no = 0 Then
''''             leave = payrs!emp_leave_type
''''          Else
''''             leave = IIf(payrs!emp_leave_period = "F", "EL", "½EL")
''''          End If
''''          sql = "update bio_device_shiftlogs set ds_status = '" & leave & "' where ds_fpcode =  '" & id & "' and ds_date =  '" & Format(payrs!emp_leave_date, "MM/dd/yyyy") & "'"
''''          paydb.Execute sql
''''          payrs.MoveNext
''''    Wend
''''    payrs.Close
''''
'''' ''for updating OD entries
''''     pst_qry = "select * from bio_emp_oddetails where empod_date between '" & Format(st_date, "MM/dd/yyyy") & "' and '" & Format(end_date, "MM/dd/yyyy") & "'"
''''    payrs.Open pst_qry, paydb, adOpenDynamic, adLockOptimistic
''''    While Not payrs.EOF
''''          id = payrs!empod_fpcode
''''          leave = "P(OD)"
''''          sql = "update bio_device_shiftlogs set ds_status = 'WOP(OD)' where ds_fpcode =  '" & id & "' and ds_date =  '" & Format(payrs!empod_date, "MM/dd/yyyy") & "' and ds_shift = 'WO'"
''''          paydb.Execute sql
''''          sql = "update bio_device_shiftlogs set ds_status = 'P(OD)' where ds_fpcode =  '" & id & "' and ds_date =  '" & Format(payrs!empod_date, "MM/dd/yyyy") & "' and ds_shift <> 'WO'"
''''          paydb.Execute sql
''''
''''          payrs.MoveNext
''''    Wend
''''    payrs.Close
''''
''''
''''
''''''for updating bio_attendlogs_daily
''''    Dim aday As String
''''    sql = "select * from bio_device_shiftlogs where ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' order by ds_date"
''''    payrs.Open sql, paydb, 1, 2
''''    While Not payrs.EOF
''''          aday = Trim(Str(Day(payrs!ds_date)))
''''          pst_qry = "update bio_attendlogs set a_day" & aday & " = '" & payrs!ds_status & "',a_in_day" & aday & " = '" & Format(payrs!ds_shift_in, "MM/dd/yyyy HH:MM:SS") & "' ,a_out_day" & aday & " = '" & Format(payrs!ds_shift_out, "MM/dd/yyyy HH:MM:SS") & "' where a_bioid = " & payrs!ds_empid & " and a_year = " & Val(cmb_year.Text) & "  and a_month = " & cmb_month.ItemData(cmb_month.ListIndex)
''''          paydb.Execute pst_qry
''''          payrs.MoveNext
''''    Wend
''''    payrs.Close
'''''end
''''
''''    Dim dayfind, dayfind_intime, dayfind_outtime As String
''''    Dim present, absent, hop, wop, cl, sl, h, ch, layoff, wo, pl As Single
''''
''''    sql = "select * from bio_attendlogs where a_year = " & Val(cmb_year.Text) & "  and a_month = " & cmb_month.ItemData(cmb_month.ListIndex)
''''    payrs.Open sql, paydb, 1, 2
''''    If Not payrs.EOF Then
''''       While Not payrs.EOF
''''            present = 0
''''            absent = 0
''''            hop = 0
''''            wop = 0
''''            cl = 0
''''            sl = 0
''''            h = 0
''''            ch = 0
''''            layoff = 0
''''            wo = 0
''''            pl = 0
''''
''''            For i = 1 To 31
''''                dayfind = "a_day" & i
''''                If payrs.Fields(dayfind) = "P" Or payrs.Fields(dayfind) = "P(OD)" Or payrs.Fields(dayfind) = "½P(OD)" Or payrs.Fields(dayfind) = "A(OD)" Then
''''                    present = present + 1
''''                ElseIf payrs.Fields(dayfind) = "A" Then
''''                    absent = absent + 1
''''                ElseIf payrs.Fields(dayfind) = "PL" Or payrs.Fields(dayfind) = "PLP" Then
''''                    pl = pl + 1
''''                ElseIf payrs.Fields(dayfind) = "½PL" Then
''''                    pl = pl + 0.5
''''                    absent = absent + 0.5
''''                ElseIf payrs.Fields(dayfind) = "½PLP" Then
''''                    pl = pl + 0.5
''''                    present = present + 0.5
''''                ElseIf payrs.Fields(dayfind) = "CL" Or payrs.Fields(dayfind) = "CL½P" Or payrs.Fields(dayfind) = "CLP" Then
''''                    cl = cl + 1
''''                ElseIf payrs.Fields(dayfind) = "½CL" Then
''''                    absent = absent + 0.5
''''                    cl = cl + 0.5
''''                ElseIf payrs.Fields(dayfind) = "½CLP" Or payrs.Fields(dayfind) = "½CL½P" Then
''''                    present = present + 0.5
''''                    cl = cl + 0.5
''''                ElseIf payrs.Fields(dayfind) = "½C.HP" Then
''''                    present = present + 0.5
''''                    ch = ch + 0.5
''''
''''                ElseIf payrs.Fields(dayfind) = "SL" Or payrs.Fields(dayfind) = "SLP" Then
''''                    sl = sl + 1
''''                ElseIf payrs.Fields(dayfind) = "½SLP" Then
''''                    sl = sl + 0.5
''''                    present = present + 0.5
''''                ElseIf payrs.Fields(dayfind) = "H" Then
''''                    h = h + 1
''''                ElseIf payrs.Fields(dayfind) = "HP" Then
''''                    hop = hop + 1
''''                ElseIf payrs.Fields(dayfind) = "½HP" Then
''''''                    hop = hop + 0.5
''''                    hop = hop + 0
''''                ElseIf payrs.Fields(dayfind) = "½P" Then
''''                    present = present + 0.5
''''                    absent = absent + 0.5
''''                ElseIf payrs.Fields(dayfind) = "Layoff" Or payrs.Fields(dayfind) = "LayoffP" Then
''''                    layoff = layoff + 1
''''                ElseIf payrs.Fields(dayfind) = "C.H" Or payrs.Fields(dayfind) = "C.H½P" Or payrs.Fields(dayfind) = "C.HP" Or payrs.Fields(dayfind) = "C.HP(OD)" Then
''''                    ch = ch + 1
''''                ElseIf payrs.Fields(dayfind) = "HOP" Or payrs.Fields(dayfind) = "H½P(OD)" Then
''''                    hop = hop + 1
''''                ElseIf payrs.Fields(dayfind) = "WOP" Or payrs.Fields(dayfind) = "WOP(OD)" Or payrs.Fields(dayfind) = "WO(OD)" Then
''''                    wop = wop + 1
''''                ElseIf payrs.Fields(dayfind) = "WO" Or payrs.Fields(dayfind) = "WO½P" Then
''''                    wo = wo + 1
''''                ElseIf payrs.Fields(dayfind) = "½C.H" Then
''''                    ch = ch + 0.5
''''                    absent = absent + 0.5
''''
''''                End If
''''            Next
''''
''''            payrs("a_present") = present
''''            payrs("a_hop") = hop
''''            payrs("a_wop") = wop
''''            payrs("a_el") = cl
''''            payrs("a_pl") = pl
''''            payrs("a_ml") = sl
''''            payrs("a_holiday") = h
''''            payrs("a_ch") = ch
''''            payrs("a_layoff") = layoff
''''            payrs("a_absent") = absent
''''            payrs("a_wo") = wo
''''
''''            payrs.Update
''''            payrs.MoveNext
''''        Wend
''''    End If
''''    payrs.Close
''''
'''' '' For reversing Holiday present if avail CH
''''    Dim holidaych As Single
''''    sql = "select * ,case when emp_ch_period = 'F' then 1 else 0.5 end as Holidaych from bio_device_shiftlogs a, bio_emp_chleave b where  ds_fpcode =  empch_fpcode and ds_shift = 'H' and ds_date = empch_worked_date and ds_date Between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' order by ds_date"
''''    payrs.Open sql, paydb, 1, 2
''''    While Not payrs.EOF
''''          idate = payrs!ds_date
''''          id = payrs!ds_fpcode
''''          holidaych = payrs!holidaych
''''          pst_qry = "update bio_attendlogs set a_hop  = a_hop - " & holidaych & "  where a_fpcode = " & id & " and a_year = " & Val(cmb_year.Text) & "  and a_month = " & cmb_month.ItemData(cmb_month.ListIndex) & " "
''''          paydb2.Execute pst_qry
''''          payrs.MoveNext
''''    Wend
''''    payrs.Close
'''''end
''''
''''
''''
''''    etime = TimeValue(Now)
''''
''''    MsgBox ("Process Start by " + Str(stime) + " ...  end by " + Str(etime))
''''
''''
''''
''''
''''
''''    MsgBox ("Updated...")
''''    Exit Sub
''''err_handler:
''''     chk = gen_Validation(Err.Number, Err.Description)
''''    '' paydb.RollbackTrans
''''     Me.MousePointer = 1
''''  '  chk = gen_Validation(Err.Number, Err.Description)
''''      If chk = 1 Then Resume
''''
''''End Sub
''''



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
    


    If end_date.Value > Now Then
       end_date.Value = Now + 1
    End If
End Sub


''Private Sub Timer1_Timer()
''    If ProgressBar1.Value + 10 > ProgressBar1.Max Then
''         ProgressBar1.Max = ProgressBar1.Max + 100
''    End If
''    ProgressBar1.Value = (ProgressBar1.Value + 1)
''End Sub
Private Sub save_Click()

End Sub
