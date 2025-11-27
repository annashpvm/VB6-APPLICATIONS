VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frm_bio_process 
   Caption         =   "REPROCESS for Employeewise"
   ClientHeight    =   8880
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8880
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   10200
      TabIndex        =   31
      Top             =   7200
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Frame Frame9 
      Caption         =   "EMP Type"
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
      Height          =   1215
      Left            =   9720
      TabIndex        =   28
      Top             =   240
      Visible         =   0   'False
      Width           =   1695
      Begin VB.OptionButton opt_cs 
         Caption         =   "Casual"
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
         Left            =   240
         TabIndex        =   32
         Top             =   720
         Width           =   1335
      End
      Begin VB.OptionButton opt_vou 
         Caption         =   "Voucher"
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
         Left            =   240
         TabIndex        =   30
         Top             =   480
         Width           =   1335
      End
      Begin VB.OptionButton opt_regular 
         Caption         =   "Regular"
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
         Left            =   240
         TabIndex        =   29
         Top             =   240
         Value           =   -1  'True
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      Height          =   5295
      Left            =   3720
      TabIndex        =   25
      Top             =   1440
      Width           =   7695
      Begin VB.CommandButton cmd_process 
         Caption         =   "RE-CALCULATE"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   2760
         TabIndex        =   27
         Top             =   4440
         Width           =   2535
      End
      Begin MSComctlLib.ListView lst_view 
         Height          =   4095
         Left            =   360
         TabIndex        =   26
         Top             =   240
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   7223
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
   End
   Begin VB.Frame Frame1 
      Height          =   5415
      Left            =   480
      TabIndex        =   13
      Top             =   1320
      Width           =   3015
      Begin VB.CommandButton cmd_clear 
         Caption         =   "CLEAR"
         Height          =   255
         Left            =   1440
         TabIndex        =   19
         Top             =   4920
         Width           =   975
      End
      Begin VB.CommandButton cmd_filter 
         Caption         =   "FILTER"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   4920
         Width           =   975
      End
      Begin VB.ListBox lst_employee 
         Height          =   1425
         Left            =   120
         TabIndex        =   17
         Top             =   3360
         Width           =   2655
      End
      Begin VB.ListBox lst_dept 
         Height          =   1425
         Left            =   120
         TabIndex        =   16
         Top             =   1560
         Width           =   2655
      End
      Begin VB.TextBox txt_empcode 
         Height          =   285
         Left            =   1440
         TabIndex        =   15
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox txt_empname 
         Height          =   285
         Left            =   120
         TabIndex        =   14
         Top             =   960
         Width           =   2655
      End
      Begin VB.Label Label2 
         Caption         =   "Employee"
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
         Left            =   120
         TabIndex        =   23
         Top             =   3120
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Department"
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
         Left            =   120
         TabIndex        =   22
         Top             =   1320
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
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Emp.Name"
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
         Index           =   3
         Left            =   120
         TabIndex        =   20
         Top             =   720
         Width           =   1095
      End
   End
   Begin VB.Frame Frame5 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   4560
      TabIndex        =   10
      Top             =   7080
      Width           =   2175
      Begin VB.CommandButton exit 
         BackColor       =   &H00C0E0FF&
         Caption         =   "&Exit"
         Height          =   705
         Left            =   1080
         MaskColor       =   &H000000FF&
         Picture         =   "frm_bio_process.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton Refresh 
         BackColor       =   &H00C0E0FF&
         Caption         =   "&Refresh"
         Height          =   705
         Left            =   120
         MaskColor       =   &H000000FF&
         Picture         =   "frm_bio_process.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   120
         Width           =   735
      End
   End
   Begin VB.Frame Frame6 
      Height          =   615
      Left            =   600
      TabIndex        =   5
      Top             =   840
      Width           =   6255
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
         TabIndex        =   7
         Top             =   120
         Width           =   3015
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
         Left            =   4680
         TabIndex        =   6
         Top             =   120
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "YEAR"
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
         Height          =   285
         Index           =   9
         Left            =   3960
         TabIndex        =   9
         Top             =   240
         Width           =   555
      End
      Begin VB.Label Label4 
         Caption         =   "MONTH"
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
         Height          =   210
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame7 
      Height          =   975
      Left            =   360
      TabIndex        =   0
      Top             =   6720
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
         Format          =   130154497
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
         Format          =   130154497
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
         TabIndex        =   4
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
         TabIndex        =   3
         Top             =   240
         Width           =   1095
      End
   End
   Begin Crystal.CrystalReport cry_rep1 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin MSComCtl2.DTPicker month_start_date 
      Height          =   375
      Left            =   240
      TabIndex        =   33
      Top             =   7920
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   0   'False
      Format          =   130154497
      CurrentDate     =   44565
   End
   Begin VB.Label lbl_emp 
      Alignment       =   2  'Center
      Caption         =   "EMPLOYEE ATTENDANCE REPROCESS"
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
      TabIndex        =   24
      Top             =   240
      Width           =   10695
   End
End
Attribute VB_Name = "frm_bio_process"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mdays As Integer
Dim del_leave As Integer
Private Sub cmd_clear_Click()
    Dim i As Long
    For i = 0 To lst_employee.ListCount - 1
        lst_employee.Selected(i) = False
    Next
End Sub

Private Sub cmd_filter_Click()
    Dim chk As Integer
    chk = 0
    Refresh_Click
    Dim payrs As New ADODB.Recordset
''    Dim itmX As ListItem
    Dim itmX As MSComctlLib.ListItem
    lst_view.ColumnHeaders.Clear
    lst_view.ColumnHeaders.Add , , "FP Code ", 1000
    lst_view.ColumnHeaders.Add , , "Emp. Name ", 2000
    lst_view.ColumnHeaders.Add , , "Department ", 1500
    lst_view.View = lvwReport
    lst_view.ListItems.Clear
    
    If txt_empcode.Text <> "" Then
      sql = "select * from bio_empmas where bioemp_fpcode =  '" & txt_empcode.Text & "' and bioemp_status = 'Working' order by bioemp_dept"
    ElseIf txt_empname.Text <> "" Then
      sql = "select * from bio_empmas where bioemp_name like  '%" & txt_empname.Text & "%' and bioemp_status = 'Working' order by bioemp_dept"
    ElseIf lst_employee.Text <> "" Then
          sql = "select * from bio_empmas where bioemp_name =  '" & lst_employee.Text & "' and bioemp_status = 'Working' order by bioemp_dept"
    Else
          sql = "select * from bio_empmas where bioemp_dept =  '" & lst_dept.Text & "' and bioemp_status = 'Working' order by bioemp_name"
    End If
    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
    While Not payrs.EOF()
            Set itmX = lst_view.ListItems.Add(, , CStr(payrs("bioemp_fpcode")))
            itmX.SubItems(1) = payrs.Fields("bioemp_name")
            itmX.SubItems(2) = payrs.Fields("bioemp_dept")
''            If txt_empcode.Text <> "" Then
            If chk = 0 Then
               itmX.Checked = True
            End If
            chk = 1
            payrs.MoveNext
    Wend
    payrs.Close
    
    
End Sub

Private Sub cmd_process_Click()
On Error GoTo err_handler
    
    
    
    Dim id, fcode As Integer
    Dim dlogdate As Date
    
    Dim dev_log(100) As Long
    
    Dim log_details As String
    Dim dev_date(100) As Date
    Dim codelist, sel_codes As String
    Dim date_details As String
    
    Dim selectcount As Integer
    
    Dim iSelected As Integer
    Dim item As ListItem
    For i = 1 To lst_view.ListItems.Count
        If lst_view.ListItems(i).Checked = True Then
          iSelected = iSelected + 1
        End If
    Next
    If iSelected = 0 Then
       MsgBox ("Employee Not selected in the view...")
       Exit Sub
    End If
    sel_codes = ""
    codelist = "(10"
    Dim idate As Date
    For i = 1 To lst_view.ListItems.Count
        If lst_view.ListItems(i).Checked = True Then
          codelist = codelist + ", " + lst_view.ListItems(i).Text
          If sel_codes = "" Then
             sel_codes = "{bio_attendlogs.a_fpcode} = " & lst_view.ListItems(i).Text
          Else
             sel_codes = sel_codes + " or {bio_attendlogs.a_fpcode} = " & lst_view.ListItems(i).Text
          End If
        End If
    Next
    codelist = codelist + ")"
''04012024
''    If lst_view.ListItems(1).Text = "3402" Then
''
''       pst_qry = "delete from bio_device_shiftlogs where ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and ds_fpcode = 3402"
''    paydb.Execute pst_qry
''
''        For idate = st_date To end_date
''             pst_qry = "insert into  bio_device_shiftlogs (ds_fpcode,ds_empid,ds_date,ds_month,ds_year) values (  '3402', '3402',  '" & Format(idate, "MM/dd/yyyy") & "', " & cmb_month.ItemData(cmb_month.ListIndex) & " , " & Val(cmb_year.Text) & "   )"
''             paydb.Execute pst_qry
''        Next
''
''    End If

    
    
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
    
   
    
    Dim dsnmdb As String
    Dim mdbrs As New ADODB.Recordset
    
    Set paydb = New ADODB.Connection
    Set payrs = New ADODB.Recordset
    
    paydb.Open pay

    
    pst_qry = "update bio_devicelogs set ad_upd ='N' where ad_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date + 1, "MM/dd/yyyy") & "' and ad_fpcode in " & codelist
    paydb.Execute pst_qry


    pst_qry = "update bio_devicelogs set ad_upd ='Y' where ad_punch = 'out'  AND ad_auto = 'A' and ad_date = '" & Format(month_start_date.Value, "MM/dd/yyyy") & "' and DATEPART(HOUR, ad_logdate) < 8 and ad_fpcode in " & codelist
    paydb.Execute pst_qry


    sql = "update bio_device_shiftlogs set ds_shift_original = '', ds_shift = 'GS',ds_shift_actual = 'GS',ds_status = '', ds_no_of_punches = 0, ds_shift_in = 0 , ds_shift_out = 0,ds_shift_in2 = 0 , ds_shift_out2 = 0,ds_shift_in3 = 0 , ds_shift_out3 = 0,ds_shift_in4 = 0 , ds_shift_out4 = 0, ds_shift_in5 = 0 , ds_shift_out5 = 0 ,ds_shift_in6 = 0 , ds_shift_out6 = 0, ds_per_hrs = 0,ds_od_hrs = 0, ds_sft_hrs = 0,ds_sft_hrs1 = 0,ds_sft_hrs2 = 0  ,ds_sft_hrs3 = 0,ds_sft_hrs4 = 0,ds_sft_hrs5 = 0  ,ds_sft_hrs6 = 0   where ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and  ds_fpcode in " & codelist
    paydb.Execute sql


    
''                For idate = st_date To end_date
''
''                   pst_qry = "insert into  bio_device_shiftlogs (ds_fpcode,ds_empid,ds_date,ds_month,ds_year) values (  '3333', '3333',  '" & Format(idate, "MM/dd/yyyy") & "', " & cmb_month.ItemData(cmb_month.ListIndex) & " , " & Val(cmb_year.Text) & " )"
''                   paydb.Execute pst_qry
''
''                   pst_qry = "insert into  bio_device_shiftlogs (ds_fpcode,ds_empid,ds_date,ds_month,ds_year) values (  '3334', '3334',  '" & Format(idate, "MM/dd/yyyy") & "', " & cmb_month.ItemData(cmb_month.ListIndex) & " , " & Val(cmb_year.Text) & " )"
''                   paydb.Execute pst_qry


'''' insert into bio_attendlogs (a_bioid,a_fpcode,a_month,a_year)  values(3333,3333,11,2022)
''
''                Next
                
    
    pst_qry = "select *  from bio_shift_schedule where emps_date  between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "'  and emps_fpcode in " & codelist
    payrs.Open pst_qry, paydb, adOpenDynamic, adLockOptimistic
    While Not payrs.EOF
          id = payrs!emps_fpcode
          sft = payrs!emps_shift
          sft_from_date = payrs!emps_date
          
''          sql = "update bio_device_shiftlogs set ds_shift = '" & sft & "',ds_shift_begintime = '" & sft_bt & "',ds_shift_endtime = '" & sft_et & "',ds_begin_duration  = '" & sft_begin_dur & "',ds_end_duration = '" & sft_begin_dur & "' where ds_empid =  '" & id & "' and ds_date between '" & Format(sft_from_date, "MM/dd/yyyy") & "' and '" & Format(sft_end_date, "MM/dd/yyyy") & "' "
          sql = "update bio_device_shiftlogs set ds_shift = '" & sft & "',ds_shift_actual = '" & sft & "' where ds_fpcode =  '" & id & "' and ds_date =  '" & Format(sft_from_date, "MM/dd/yyyy") & "'"

          paydb.Execute sql
          payrs.MoveNext
    Wend
    payrs.Close


        Set paydb2 = New ADODB.Connection
        Set payrs2 = New ADODB.Recordset
        paydb2.Open pay
        
        Dim j As Integer
        Dim dt_log As Integer
        Dim inpunch_chk, outpunch_chk, incount, outcount   As Integer


       Dim rs_set As New ADODB.Recordset
       Dim newqry As String
       
       
       Dim inoutchk As Integer
       Dim logdate As Date
    ''Annadurai
 '' end_date = end_date - 1
        sql = "select * from bio_device_shiftlogs where ds_fpcode = 1018 and  ds_date  between  '" & Format(st_date.Value, "MM/dd/yyyy") & "' and '" & Format(end_date.Value, "MM/dd/yyyy") & "' and ds_fpcode > 0  order by ds_fpcode,ds_date"
        newqry = "select * from bio_device_shiftlogs where ds_date  between  '" & Format(st_date.Value, "MM/dd/yyyy") & "' and '" & Format(end_date.Value, "MM/dd/yyyy") & "' and ds_fpcode > 0  order by ds_fpcode,ds_date"
       
        newqry = "select ds_fpcode,ds_date from bio_device_shiftlogs,bio_devicelogs where ds_fpcode = ad_fpcode and ds_date = ad_date and ds_date  between  '" & Format(st_date.Value, "MM/dd/yyyy") & "' and '" & Format(end_date.Value, "MM/dd/yyyy") & "' and ds_fpcode > 0   and ds_fpcode in " & codelist & " group by ds_fpcode,ds_date order by ds_fpcode,ds_date"
        
    ''    newqry = "select ds_fpcode,ds_date from bio_device_shiftlogs,bio_devicelogs where ds_fpcode = 10006 and ds_fpcode = ad_fpcode  and ds_date  between  '" & Format(st_date.Value, "MM/dd/yyyy") & "' and '" & Format(end_date.Value, "MM/dd/yyyy") & "' and ds_fpcode > 0  group by ds_fpcode,ds_date order by ds_fpcode,ds_date"
        
        rs_set.Open newqry, paydb, 1, 2
        While Not rs_set.EOF
                 id = rs_set!ds_fpcode
              idate = rs_set!ds_date
              sql = "select * from bio_device_shiftlogs where ds_fpcode = " & id & " and ds_date = '" & Format(idate, "MM/dd/yyyy") & "' "
              payrs.Open sql, paydb, 1, 2
              While Not payrs.EOF
              dt_log = 1
              i = 1
              
''              MsgBox (payrs!ds_date)

              id = payrs!ds_fpcode
              sft_from_date = payrs!ds_date
''              If sft_from_date = "17/10/2025" Then
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
        


        
        
150:
    paydb.CommandTimeout = 300
''    sql = "update bio_device_shiftlogs set ds_shift = 'H',ds_shift_actual = 'H',ds_status = 'H',ds_shift_begintime = '00:00',ds_shift_endtime = '00:00',ds_begin_duration  = '',ds_end_duration = '' ,ds_sft_hrs = 0  from bio_device_shiftlogs a, emp_dec_holiday_empwise b where emp_decholi_date = ds_date and  emp_decholi_fpcode = ds_fpcode and  emp_decholi_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "'  and ds_fpcode in " & codelist
    
    

    sql = "update bio_device_shiftlogs set ds_shift = 'H',ds_shift_actual = 'H',ds_status = 'H',ds_shift_begintime = '00:00',ds_shift_endtime = '00:00',ds_begin_duration  = '',ds_end_duration = '' ,ds_sft_hrs = 0  from bio_device_shiftlogs a, emp_dec_holiday b , emp_mas c   where emp_dec_holiday = ds_date  and emp_dec_holiday Between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and ds_date >= emp_doj and  ds_fpcode = emp_fpcode   and ds_fpcode in " & codelist
    
    paydb.Execute sql
    
    
    paydb.Execute sql
    
    
''    Dim woday As String
    pst_qry = "select * from emp_mas where emp_company in (1) and emp_status = 'A' and emp_fpcode in " & codelist
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
''         If payrs("emp_cat") = "W" Then
''            sql = "update bio_device_shiftlogs set ds_shift_original = 'WOH', ds_status = 'WOH', ds_shift = 'WOH',ds_shift_actual = 'WOH',ds_shift_begintime = '',ds_shift_endtime = '',ds_begin_duration  = '',ds_end_duration = '' where ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and ds_fpcode = " & payrs("emp_fpcode") & " and ds_shift = 'H' and datepart(dw,ds_date) =  " & woday
''            sql = "update bio_device_shiftlogs set ds_shift_original = 'WOH', ds_status = 'WOH', ds_shift = 'WOH',ds_shift_begintime = '',ds_shift_endtime = '',ds_begin_duration  = '',ds_end_duration = '' where ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and ds_fpcode = " & payrs("emp_fpcode") & " and ds_shift = 'H' and datepart(dw,ds_date) =  " & woday
''            paydb.Execute sql
''            sql = "update bio_device_shiftlogs set ds_shift_original = 'WO', ds_status = 'WO', ds_shift = 'WO',ds_shift_actual = 'WO',ds_shift_begintime = '',ds_shift_endtime = '',ds_begin_duration  = '',ds_end_duration = '' where ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and ds_fpcode = " & payrs("emp_fpcode") & " and ds_shift <> 'WOH' and datepart(dw,ds_date) =  " & woday
''            sql = "update bio_device_shiftlogs set ds_shift_original = 'WO', ds_status = 'WO', ds_shift = 'WO',ds_shift_begintime = '',ds_shift_endtime = '',ds_begin_duration  = '',ds_end_duration = '' where ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and ds_fpcode = " & payrs("emp_fpcode") & " and ds_shift <> 'WOH' and datepart(dw,ds_date) =  " & woday
''            paydb.Execute sql
''
''         Else
''            sql = "update bio_device_shiftlogs set ds_shift_original = 'WO', ds_status = 'WO', ds_shift = 'WO',ds_shift_actual = 'WO',ds_shift_begintime = '',ds_shift_endtime = '',ds_begin_duration  = '',ds_end_duration = '' where ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and ds_fpcode = " & payrs("emp_fpcode") & " and datepart(dw,ds_date) =  " & woday
''            sql = "update bio_device_shiftlogs set ds_shift_original = 'WO', ds_status = 'WO', ds_shift = 'WO',ds_shift_begintime = '',ds_shift_endtime = '',ds_begin_duration  = '',ds_end_duration = '' where ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and ds_fpcode = " & payrs("emp_fpcode") & " and datepart(dw,ds_date) =  " & woday
''            paydb.Execute sql
''         End If
''

         If payrs("emp_cat") = "W" Then
            sql = "update bio_device_shiftlogs set ds_shift_original = 'WOH', ds_status = 'WOH', ds_shift = 'WOH',ds_shift_begintime = '',ds_shift_endtime = '',ds_begin_duration  = '',ds_end_duration = '' from bio_device_shiftlogs  , emp_mas     where ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and ds_fpcode = " & payrs("emp_fpcode") & " and ds_shift = 'H' and ds_date >= emp_doj and  ds_fpcode = emp_fpcode    and datepart(dw,ds_date) =  " & woday
            paydb.Execute sql
            sql = "update bio_device_shiftlogs set ds_shift_original = 'WO', ds_status = 'WO', ds_shift = 'WO',ds_shift_begintime = '',ds_shift_endtime = '',ds_begin_duration  = '',ds_end_duration = '' from bio_device_shiftlogs  , emp_mas      where ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and ds_fpcode = " & payrs("emp_fpcode") & " and ds_shift <> 'WOH' and ds_date >= emp_doj and  ds_fpcode = emp_fpcode   and datepart(dw,ds_date) =  " & woday
            paydb.Execute sql
         
         Else
            sql = "update bio_device_shiftlogs set ds_shift_original = 'WO', ds_status = 'WO', ds_shift = 'WO',ds_shift_begintime = '',ds_shift_endtime = '',ds_begin_duration  = '',ds_end_duration = '' from bio_device_shiftlogs  , emp_mas     where ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and ds_fpcode = " & payrs("emp_fpcode") & " and ds_date >= emp_doj and  ds_fpcode = emp_fpcode    and datepart(dw,ds_date) =  " & woday
            paydb.Execute sql
         End If
         

         payrs.MoveNext
    Wend
    payrs.Close


    
''updating voucher payment
    pst_qry = "select * from emp_voupay_mast where emp_company in (1,2,3,5)  and emp_cat <> 'M' and emp_fpcode in " & codelist
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
100:
    
        sql = "update bio_device_shiftlogs  set ds_sft_hrs1=  (datediff(minute,ds_shift_in, ds_shift_out)/60)+convert(decimal,(datediff(minute,ds_shift_in, ds_shift_out) -(datediff(minute,ds_shift_in, ds_shift_out)/60 * 60)))/100 where ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and ds_shift_out > ds_shift_in and ds_shift_in > 0  and ds_fpcode in " & codelist
        paydb.Execute sql
        sql = "update bio_device_shiftlogs  set ds_sft_hrs2 = (datediff(minute,ds_shift_in2, ds_shift_out2)/60)+convert(decimal,(datediff(minute,ds_shift_in2, ds_shift_out2) -(datediff(minute,ds_shift_in2, ds_shift_out2)/60 * 60)))/100 where ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and ds_shift_out2 > ds_shift_in2 and ds_shift_in2 > 0  and ds_fpcode in " & codelist
        paydb.Execute sql
        sql = "update bio_device_shiftlogs  set ds_sft_hrs3 = (datediff(minute,ds_shift_in3, ds_shift_out3)/60)+convert(decimal,(datediff(minute,ds_shift_in3, ds_shift_out3) -(datediff(minute,ds_shift_in3, ds_shift_out3)/60 * 60)))/100 where ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and ds_shift_out3 > ds_shift_in3 and ds_shift_in3 > 0  and ds_fpcode in " & codelist
        paydb.Execute sql
        sql = "update bio_device_shiftlogs  set ds_sft_hrs4 = (datediff(minute,ds_shift_in4, ds_shift_out4)/60)+convert(decimal,(datediff(minute,ds_shift_in4, ds_shift_out4) -(datediff(minute,ds_shift_in4, ds_shift_out4)/60 * 60)))/100 where ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and ds_shift_out4 > ds_shift_in4 and ds_shift_in4 > 0  and ds_fpcode in " & codelist
        paydb.Execute sql
        sql = "update bio_device_shiftlogs  set ds_sft_hrs5 = (datediff(minute,ds_shift_in5, ds_shift_out5)/60)+convert(decimal,(datediff(minute,ds_shift_in5, ds_shift_out5) -(datediff(minute,ds_shift_in5, ds_shift_out5)/60 * 60)))/100 where ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and ds_shift_out5 > ds_shift_in5 and ds_shift_in5 > 0  and ds_fpcode in " & codelist
        paydb.Execute sql
        sql = "update bio_device_shiftlogs  set ds_sft_hrs6 = (datediff(minute,ds_shift_in6, ds_shift_out6)/60)+convert(decimal,(datediff(minute,ds_shift_in6, ds_shift_out6) -(datediff(minute,ds_shift_in6, ds_shift_out6)/60 * 60)))/100 where ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and ds_shift_out6 > ds_shift_in6 and ds_shift_in6 > 0  and ds_fpcode in " & codelist
        paydb.Execute sql

    

    
    
    Dim firsttime, endtime As Double
    Dim per_totmins, per_hrs, per_mins, per_time As Double
 ''for updating Permission entries
    pst_qry = "select * from bio_emp_permissions where empp_date between '" & Format(st_date, "MM/dd/yyyy") & "' and '" & Format(end_date, "MM/dd/yyyy") & "' and empp_fpcode in " & codelist
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

    pst_qry = "select * from bio_emp_oddetails where empod_date between '" & Format(st_date, "MM/dd/yyyy") & "' and '" & Format(end_date, "MM/dd/yyyy") & "'  and empod_fpcode in " & codelist
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
    sql = "select * from bio_device_shiftlogs where ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and '" & Format(end_date, "MM/dd/yyyy") & "'  and ds_sft_hrs1+ds_sft_hrs2+ds_sft_hrs3+ds_sft_hrs4+ds_sft_hrs5+ds_sft_hrs6+ds_per_hrs+ds_od_hrs >0 and ds_fpcode in " & codelist
    payrs.Open sql, paydb, 1, 2
    While Not payrs.EOF
          
          wmins1 = Int(Val(payrs!ds_sft_hrs1)) * 60 + (Val(payrs!ds_sft_hrs1) - Int(Val(payrs!ds_sft_hrs1))) * 100
          wmins2 = Int(Val(payrs!ds_sft_hrs2)) * 60 + (Val(payrs!ds_sft_hrs2) - Int(Val(payrs!ds_sft_hrs2))) * 100
          wmins3 = Int(Val(payrs!ds_sft_hrs3)) * 60 + (Val(payrs!ds_sft_hrs3) - Int(Val(payrs!ds_sft_hrs3))) * 100
          wmins4 = Int(Val(payrs!ds_sft_hrs4)) * 60 + (Val(payrs!ds_sft_hrs4) - Int(Val(payrs!ds_sft_hrs4))) * 100
          wmins5 = Int(Val(payrs!ds_sft_hrs5)) * 60 + (Val(payrs!ds_sft_hrs5) - Int(Val(payrs!ds_sft_hrs5))) * 100
          wmins6 = Int(Val(payrs!ds_sft_hrs6)) * 60 + (Val(payrs!ds_sft_hrs6) - Int(Val(payrs!ds_sft_hrs6))) * 100
          permins = Int(Val(payrs!ds_per_hrs)) * 60 + (Val(payrs!ds_per_hrs) - Int(Val(payrs!ds_per_hrs))) * 100
''          If payrs!ds_per_hrs > 0 Then
''             MsgBox ("Wait")
''          End If
          
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
 
 

    
''    sql = "update bio_device_shiftlogs set ds_status = 'A' WHERE ds_sft_hrs = 0 and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(Now - 1, "MM/dd/yyyy") & "' and ds_shift <> 'WO'  and ds_shift <> 'WOH' and ds_fpcode in " & codelist
''    paydb.Execute sql
''
''    sql = "update bio_device_shiftlogs set ds_status = 'A' WHERE  ds_sft_hrs < 3 and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(Now - 1, "MM/dd/yyyy") & "' and ds_shift <> 'WO' and ds_shift <> 'WOH' and ds_fpcode in " & codelist
''    paydb.Execute sql

    sql = "update bio_device_shiftlogs set ds_status = 'A' WHERE ds_sft_hrs = 0 and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(Now - 1, "MM/dd/yyyy") & "' and ds_shift <> 'WO'  and ds_shift <> 'WOH' and ds_fpcode in " & codelist
    sql = "update bio_device_shiftlogs set ds_status = 'A'  from bio_device_shiftlogs a, emp_mas b   where ds_sft_hrs = 0 and ds_date Between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and ds_date >= emp_doj and  ds_fpcode = emp_fpcode  and ds_shift <> 'WO'  and ds_shift <> 'WOH'  and ds_fpcode in " & codelist
    paydb.Execute sql
    
    sql = "update bio_device_shiftlogs set ds_status = 'A' WHERE  ds_sft_hrs < 3 and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(Now - 1, "MM/dd/yyyy") & "' and ds_shift <> 'WO' and ds_shift <> 'WOH' and ds_fpcode in " & codelist
    sql = "update bio_device_shiftlogs set ds_status = 'A'  from bio_device_shiftlogs a, emp_mas b   where ds_sft_hrs <3 and ds_date Between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and ds_date >= emp_doj and  ds_fpcode = emp_fpcode  and ds_shift <> 'WO'  and ds_shift <> 'WOH'  and ds_fpcode in " & codelist
    paydb.Execute sql


    
    
210:

''New Addition for A SHIFT & B SHIFT 21/07/2020
''Start
    
    
    sql = "update bio_device_shiftlogs set ds_shift_actual = '06.00PM-06.00AM'  from bio_device_shiftlogs WHERE  ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and '" & Format(end_date, "MM/dd/yyyy") & "' and DATEpart(hour,ds_shift_in) >= 17 and DATEpart(hour,ds_shift_in) <= 18 and ds_fpcode in " & codelist
    paydb.Execute sql
    sql = "update bio_device_shiftlogs set ds_shift_actual = '07.00PM-07.00AM'  from bio_device_shiftlogs WHERE  ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and '" & Format(end_date, "MM/dd/yyyy") & "' and DATEpart(hour,ds_shift_in) >= 18 and DATEpart(hour,ds_shift_in) <= 19 and ds_fpcode in " & codelist
    paydb.Execute sql
    sql = "update bio_device_shiftlogs set ds_shift_actual = '08.00PM-08.00AM'  from bio_device_shiftlogs WHERE  ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and '" & Format(end_date, "MM/dd/yyyy") & "' and DATEpart(hour,ds_shift_in) >= 19 and DATEpart(hour,ds_shift_in) <= 20 and ds_fpcode in " & codelist
    paydb.Execute sql
    
    sql = "update bio_device_shiftlogs set ds_shift_actual = 'A SHIFT'  from bio_device_shiftlogs WHERE  ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and DATEpart(hour,ds_shift_in) >= 5 and DATEpart(hour,ds_shift_in) <= 6 and   DATEpart(hour,ds_shift_out) >= 13  and  DATEpart(hour,ds_shift_out) <= 16 and ds_fpcode in " & codelist
    paydb.Execute sql
    sql = "update bio_device_shiftlogs set ds_shift_actual = 'B SHIFT'  from bio_device_shiftlogs WHERE  ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and DATEpart(hour,ds_shift_in) >= 13 and DATEpart(hour,ds_shift_in) <= 16 and  DATEpart(hour,ds_shift_out) >= 21  and  DATEpart(hour,ds_shift_out) <= 23 and ds_fpcode in " & codelist
    paydb.Execute sql
    sql = "update bio_device_shiftlogs set ds_shift_actual = 'C SHIFT'  from bio_device_shiftlogs WHERE  ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and DATEpart(hour,ds_shift_in) >= 21 and DATEpart(hour,ds_shift_in) <= 22 and  DATEpart(hour,ds_shift_out) >= 5  and  DATEpart(hour,ds_shift_out) <= 10 and ds_fpcode in " & codelist
    paydb.Execute sql
    sql = "update bio_device_shiftlogs set ds_shift_actual = '06.00PM-06.00AM'  from bio_device_shiftlogs WHERE  ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and '" & Format(end_date, "MM/dd/yyyy") & "' and DATEpart(hour,ds_shift_in) >= 17 and DATEpart(hour,ds_shift_in) <= 18 and  DATEpart(hour,ds_shift_out) >= 6  and  DATEpart(hour,ds_shift_out) <= 7 and ds_fpcode in " & codelist
    paydb.Execute sql
    sql = "update bio_device_shiftlogs set ds_shift_actual = '07.00PM-07.00AM'  from bio_device_shiftlogs WHERE  ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and '" & Format(end_date, "MM/dd/yyyy") & "' and DATEpart(hour,ds_shift_in) >= 18 and DATEpart(hour,ds_shift_in) <= 19 and  DATEpart(hour,ds_shift_out) >= 7  and  DATEpart(hour,ds_shift_out) <= 8 and ds_fpcode in " & codelist
    paydb.Execute sql
    sql = "update bio_device_shiftlogs set ds_shift_actual = '08.00PM-08.00AM'  from bio_device_shiftlogs WHERE  ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and '" & Format(end_date, "MM/dd/yyyy") & "' and DATEpart(hour,ds_shift_in) >= 19 and DATEpart(hour,ds_shift_in) <= 20 and  DATEpart(hour,ds_shift_out) >= 8  and  DATEpart(hour,ds_shift_out) <= 9 and ds_fpcode in " & codelist
    paydb.Execute sql

    sql = "update bio_device_shiftlogs set ds_shift_actual = 'Unshift'  from bio_device_shiftlogs WHERE  ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and '" & Format(end_date, "MM/dd/yyyy") & "' and DATEpart(hour,ds_shift_in) >= 15 and DATEpart(hour,ds_shift_in) <= 20 and  DATEpart(hour,ds_shift_out) >= 2  and  DATEpart(hour,ds_shift_out) <= 4 and ds_fpcode in " & codelist
    paydb.Execute sql


    


''End



    sql = "update bio_device_shiftlogs set ds_status = 'H' WHERE   ds_shift = 'H' and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and ds_fpcode in " & codelist
    paydb.Execute sql

''FOR A,B & C SHIFT - COMMON FOR STAFF & WORKER - 1/2 DAY PRESENT
    sql = "update bio_device_shiftlogs set ds_status = '½P'  from bio_device_shiftlogs , bio_empmas   WHERE  ds_fpcode = bioemp_fpcode and ds_sft_hrs >3.30  and  ds_sft_hrs < 7.45 and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and ds_shift <> 'WO' and ds_shift <> 'WOH' and ds_shift_actual in ('A SHIFT','B SHIFT','C SHIFT') and ds_fpcode in " & codelist
    paydb.Execute sql
    sql = "update bio_device_shiftlogs set ds_status = 'WO½P'  from bio_device_shiftlogs , bio_empmas   WHERE  ds_fpcode = bioemp_fpcode and ds_sft_hrs >3.30  and  ds_sft_hrs < 7.45 and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and ds_shift = 'WO' and ds_shift <> 'WOH' and ds_shift_actual in ('A SHIFT','B SHIFT','C SHIFT') and ds_fpcode in " & codelist
    paydb.Execute sql
    
''FOR STAFF - 1/2 DAY PRESENT - for SINGLE IN / OUT punches

''FOR STAFF - 1/2 DAY PRESENT - for SINGLE IN / OUT punches
    sql = "update bio_device_shiftlogs set ds_status = '½P'  from bio_device_shiftlogs , bio_empmas   WHERE  ds_fpcode = bioemp_fpcode and bioemp_team = 'STAFF' and   ds_sft_hrs >3.30 and  ds_sft_hrs < 8.45 and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and ds_shift <> 'WO' and ds_shift <> 'WOH' and ds_no_of_punches = 1 and ds_shift_actual not in ('A SHIFT','B SHIFT','C SHIFT') and bioemp_gender = 'F' and ds_fpcode in " & codelist
    paydb.Execute sql
    
    sql = "update bio_device_shiftlogs set ds_status = 'WO½P'  from bio_device_shiftlogs , bio_empmas   WHERE  ds_fpcode = bioemp_fpcode and bioemp_team = 'STAFF' and  ds_sft_hrs >3.30 and  ds_sft_hrs < 8.45 and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and ds_shift = 'WO' and ds_shift <> 'WOH' and ds_no_of_punches = 1 and ds_shift_actual not in ('A SHIFT','B SHIFT','C SHIFT')  and bioemp_gender = 'F' and ds_fpcode in " & codelist
    paydb.Execute sql
    
    sql = "update bio_device_shiftlogs set ds_status = '½P'  from bio_device_shiftlogs , bio_empmas   WHERE  ds_fpcode = bioemp_fpcode and bioemp_team = 'STAFF' and   ds_sft_hrs >3.30 and  ds_sft_hrs < 8.45 and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and ds_shift <> 'WO' and ds_shift <> 'WOH' and ds_no_of_punches = 1 and ds_shift_actual not in ('A SHIFT','B SHIFT','C SHIFT') and bioemp_gender = 'M' and ds_fpcode in " & codelist
    paydb.Execute sql
    sql = "update bio_device_shiftlogs set ds_status = 'WO½P'  from bio_device_shiftlogs , bio_empmas   WHERE  ds_fpcode = bioemp_fpcode and bioemp_team = 'STAFF' and  ds_sft_hrs >3.30 and  ds_sft_hrs < 8.45 and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and ds_shift = 'WO' and ds_shift <> 'WOH' and ds_no_of_punches = 1 and ds_shift_actual not in ('A SHIFT','B SHIFT','C SHIFT')  and bioemp_gender = 'M' and ds_fpcode in " & codelist
    paydb.Execute sql
    

''FOR STAFF - 1/2 DAY PRESENT - for DOUBLE IN / OUT punches
    sql = "update bio_device_shiftlogs set ds_status = '½P'  from bio_device_shiftlogs , bio_empmas   WHERE  ds_fpcode = bioemp_fpcode and bioemp_team = 'STAFF' and   ds_sft_hrs >3.30 and  ds_sft_hrs < 7.45 and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and ds_shift <> 'WO' and ds_shift <> 'WOH' and ds_no_of_punches = 2 and ds_shift_actual not in ('A SHIFT','B SHIFT','C SHIFT') and bioemp_gender = 'F' and ds_fpcode in " & codelist
    paydb.Execute sql
    
    sql = "update bio_device_shiftlogs set ds_status = 'WO½P'  from bio_device_shiftlogs , bio_empmas   WHERE  ds_fpcode = bioemp_fpcode and bioemp_team = 'STAFF' and  ds_sft_hrs >3.30 and  ds_sft_hrs < 7.45 and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and ds_shift = 'WO' and ds_shift <> 'WOH' and ds_no_of_punches = 2 and ds_shift_actual not in ('A SHIFT','B SHIFT','C SHIFT')  and bioemp_gender = 'F' and ds_fpcode in " & codelist
    paydb.Execute sql
    
    sql = "update bio_device_shiftlogs set ds_status = '½P'  from bio_device_shiftlogs , bio_empmas   WHERE  ds_fpcode = bioemp_fpcode and bioemp_team = 'STAFF' and   ds_sft_hrs >3.30 and  ds_sft_hrs < 7.45 and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and ds_shift <> 'WO' and ds_shift <> 'WOH' and ds_no_of_punches = 2 and ds_shift_actual not in ('A SHIFT','B SHIFT','C SHIFT') and bioemp_gender = 'M' and ds_fpcode in " & codelist
    paydb.Execute sql
    sql = "update bio_device_shiftlogs set ds_status = 'WO½P'  from bio_device_shiftlogs , bio_empmas   WHERE  ds_fpcode = bioemp_fpcode and bioemp_team = 'STAFF' and  ds_sft_hrs >3.30 and  ds_sft_hrs < 7.45 and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and ds_shift = 'WO' and ds_shift <> 'WOH' and ds_no_of_punches = 2 and ds_shift_actual not in ('A SHIFT','B SHIFT','C SHIFT')  and bioemp_gender = 'M' and ds_fpcode in " & codelist
    paydb.Execute sql
    
    
    
    
''FOR UNSHIFT
''    sql = "update bio_device_shiftlogs set ds_status = '½P'  from bio_device_shiftlogs , bio_empmas   WHERE  ds_fpcode = bioemp_fpcode and bioemp_team = 'STAFF' and   ds_sft_hrs >3.30 and  ds_sft_hrs < 7.45 and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and ds_shift <> 'WO' and ds_shift <> 'WOH' and ds_no_of_punches = 1 and ds_shift_actual  in ('Unshift') and ds_fpcode in " & codelist
    sql = "update bio_device_shiftlogs set ds_status = '½P'  from bio_device_shiftlogs , bio_empmas   WHERE  ds_fpcode = bioemp_fpcode  and   ds_sft_hrs >3.30 and  ds_sft_hrs < 7.45 and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and ds_shift <> 'WO' and ds_shift <> 'WOH' and ds_no_of_punches = 1 and ds_shift_actual  in ('Unshift') and ds_fpcode in " & codelist
    paydb.Execute sql
    
   '' sql = "update bio_device_shiftlogs set ds_status = 'P'  from bio_device_shiftlogs , bio_empmas   WHERE  ds_fpcode = bioemp_fpcode and bioemp_team = 'STAFF' and   ds_sft_hrs > 7.44  and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and ds_shift <> 'WO' and ds_shift <> 'WOH' and ds_no_of_punches = 1 and ds_shift_actual  in ('Unshift') and ds_fpcode in " & codelist
    sql = "update bio_device_shiftlogs set ds_status = 'P'  from bio_device_shiftlogs , bio_empmas   WHERE  ds_fpcode = bioemp_fpcode and ds_sft_hrs > 7.44  and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and ds_shift <> 'WO' and ds_shift <> 'WOH' and ds_no_of_punches = 1 and ds_shift_actual  in ('Unshift') and ds_fpcode in " & codelist
    paydb.Execute sql
    
    
''FOR STAFF - 1/2 DAY PRESENT - for DOUBLE IN / OUT punches
    
    sql = "update bio_device_shiftlogs set ds_status = '½P'  from bio_device_shiftlogs , bio_empmas   WHERE  ds_fpcode = bioemp_fpcode and bioemp_team = 'STAFF' and   ds_sft_hrs >3.30 and  ds_sft_hrs < 7.45 and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and ds_shift <> 'WO' and ds_shift <> 'WOH' and ds_shift_actual not in ('A SHIFT','B SHIFT','C SHIFT') and bioemp_gender = 'F' and ds_fpcode in " & codelist
    paydb.Execute sql
    sql = "update bio_device_shiftlogs set ds_status = 'WO½P'  from bio_device_shiftlogs , bio_empmas   WHERE  ds_fpcode = bioemp_fpcode and bioemp_team = 'STAFF' and  ds_sft_hrs >3.30 and  ds_sft_hrs < 7.45 and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and ds_shift = 'WO' and ds_shift <> 'WOH' and ds_shift_actual not in ('A SHIFT','B SHIFT','C SHIFT') and bioemp_gender = 'F' and ds_fpcode in " & codelist
    paydb.Execute sql
    
    sql = "update bio_device_shiftlogs set ds_status = '½P'  from bio_device_shiftlogs , bio_empmas   WHERE  ds_fpcode = bioemp_fpcode and bioemp_team = 'STAFF' and   ds_sft_hrs >3.30 and  ds_sft_hrs < 8.45 and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and ds_shift <> 'WO' and ds_shift <> 'WOH' and ds_shift_actual not in ('A SHIFT','B SHIFT','C SHIFT') and bioemp_gender = 'M' and ds_fpcode in " & codelist
    paydb.Execute sql
    sql = "update bio_device_shiftlogs set ds_status = 'WO½P'  from bio_device_shiftlogs , bio_empmas   WHERE  ds_fpcode = bioemp_fpcode and bioemp_team = 'STAFF' and  ds_sft_hrs >3.30 and  ds_sft_hrs < 8.45 and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and ds_shift = 'WO' and ds_shift <> 'WOH' and ds_shift_actual not in ('A SHIFT','B SHIFT','C SHIFT') and bioemp_gender = 'M' and ds_fpcode in " & codelist
    paydb.Execute sql
    
    
''FOR DECLARE HOLIDAY - STAFF - 1/2 DAY PRESENT
    sql = "update bio_device_shiftlogs set ds_status = '½HP'  from bio_device_shiftlogs , bio_empmas   WHERE  ds_fpcode = bioemp_fpcode and bioemp_team = 'STAFF' and  ds_sft_hrs > 3.30 and  ds_sft_hrs < 7.40 and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and  ds_shift = 'H' and ds_fpcode in " & codelist
    paydb.Execute sql
    sql = "update bio_device_shiftlogs set ds_status = 'HP' from bio_device_shiftlogs , bio_empmas    WHERE ds_fpcode = bioemp_fpcode and bioemp_team = 'STAFF' and ds_sft_hrs >= 7.40 and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and  ds_shift = 'H' and ds_fpcode in " & codelist
    paydb.Execute sql
    
    
''FOR WORKER - 1/2 DAY PRESENT
    

    sql = "update bio_device_shiftlogs set ds_status = '½P'  from bio_device_shiftlogs , bio_empmas   WHERE  ds_fpcode = bioemp_fpcode and bioemp_team <> 'STAFF' and   ds_sft_hrs >3.30 and  ds_sft_hrs < 7.46 and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and ds_shift <> 'WO' and ds_shift <> 'WOH'  and ds_fpcode in " & codelist
    paydb.Execute sql
    sql = "update bio_device_shiftlogs set ds_status = 'WO½P'  from bio_device_shiftlogs , bio_empmas   WHERE  ds_fpcode = bioemp_fpcode and bioemp_team <> 'STAFF' and  ds_sft_hrs >3.30 and  ds_sft_hrs < 7.46 and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and ds_shift = 'WO' and ds_shift <> 'WOH'  and ds_fpcode in " & codelist
    paydb.Execute sql
    sql = "update bio_device_shiftlogs set ds_status = '½HP'  from bio_device_shiftlogs , bio_empmas   WHERE  ds_fpcode = bioemp_fpcode and bioemp_team <> 'STAFF' and  ds_sft_hrs > 3.30 and  ds_sft_hrs < 7.46 and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and  ds_shift = 'H' and ds_fpcode in " & codelist
    paydb.Execute sql
    sql = "update bio_device_shiftlogs set ds_status = 'HP' from bio_device_shiftlogs , bio_empmas    WHERE ds_fpcode = bioemp_fpcode and bioemp_team <> 'STAFF' and ds_sft_hrs >= 7.46 and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and  ds_shift = 'H' and ds_fpcode in " & codelist
    paydb.Execute sql
   
    
    
''FOR B & C SHIFT - COMMON FOR STAFF & WORKER - FULL DAY PRESENT
    sql = "update bio_device_shiftlogs set ds_status = 'P' from bio_device_shiftlogs , bio_empmas  WHERE ds_fpcode = bioemp_fpcode and ds_sft_hrs >= 7.45 and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "'  and  ds_shift <> 'H'  and  ds_shift <> 'WOH'  and ds_shift_actual in ('A SHIFT','B SHIFT','C SHIFT') and ds_fpcode in " & codelist
    paydb.Execute sql
    sql = "update bio_device_shiftlogs set ds_status = 'WOP' from bio_device_shiftlogs , bio_empmas  WHERE ds_fpcode = bioemp_fpcode and  ds_sft_hrs >= 7.45 and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and  ds_shift = 'WO' and ds_status = 'P'  and ds_shift_actual in ('A SHIFT','B SHIFT','C SHIFT') and ds_fpcode in " & codelist
    paydb.Execute sql
    

    '' for  STAFF - FEMALE
    sql = "update bio_device_shiftlogs set ds_status = 'P' from bio_device_shiftlogs , bio_empmas  WHERE ds_fpcode = bioemp_fpcode and bioemp_team = 'STAFF' and ds_sft_hrs >= 7.30 and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "'  and  ds_shift <> 'H'  and  ds_shift <> 'WOH' and ds_no_of_punches = 1 and ds_shift_actual not in ('A SHIFT','B SHIFT','C SHIFT')  and bioemp_gender = 'F' and ds_fpcode in " & codelist
    paydb.Execute sql
    sql = "update bio_device_shiftlogs set ds_status = 'WOP' from bio_device_shiftlogs , bio_empmas  WHERE ds_fpcode = bioemp_fpcode and bioemp_team = 'STAFF' and  ds_sft_hrs >= 7.30 and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and  ds_shift = 'WO' and ds_status = 'P' and ds_no_of_punches = 1   and ds_shift_actual not in ('A SHIFT','B SHIFT','C SHIFT')  and bioemp_gender = 'F' and ds_fpcode in " & codelist
    paydb.Execute sql
    
    sql = "update bio_device_shiftlogs set ds_status = 'WOHP' from bio_device_shiftlogs , bio_empmas  WHERE ds_fpcode = bioemp_fpcode and bioemp_team = 'STAFF' and ds_sft_hrs >= 7.30 and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and  ds_status = 'WOH'  and ds_no_of_punches = 1 and ds_shift_actual not in ('A SHIFT','B SHIFT','C SHIFT')  and bioemp_gender = 'F' and ds_fpcode in " & codelist
    paydb.Execute sql
    sql = "update bio_device_shiftlogs set ds_status = 'P' from bio_device_shiftlogs , bio_empmas WHERE ds_fpcode = bioemp_fpcode and bioemp_team = 'STAFF' and ds_sft_hrs >= 7.30 and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "'  and  ds_shift <> 'H' and  ds_shift <> 'WOH'  and  ds_shift <> 'WO' and ds_no_of_punches = 1  and ds_shift_actual not in ('A SHIFT','B SHIFT','C SHIFT')  and bioemp_gender = 'F' and ds_fpcode in " & codelist
    paydb.Execute sql
    
    sql = "update bio_device_shiftlogs set ds_status = 'P' from bio_device_shiftlogs , bio_empmas  WHERE ds_fpcode = bioemp_fpcode and bioemp_team = 'STAFF' and ds_sft_hrs >= 7.30 and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "'  and  ds_shift <> 'H'  and  ds_shift <> 'WOH' and ds_shift_actual not in ('A SHIFT','B SHIFT','C SHIFT')  and bioemp_gender = 'F' and ds_fpcode in " & codelist
    paydb.Execute sql
    sql = "update bio_device_shiftlogs set ds_status = 'WOP' from bio_device_shiftlogs , bio_empmas  WHERE ds_fpcode = bioemp_fpcode and bioemp_team = 'STAFF' and  ds_sft_hrs >= 7.30 and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and  ds_shift = 'WO' and ds_status = 'P'   and ds_shift_actual not in ('A SHIFT','B SHIFT','C SHIFT')  and bioemp_gender = 'F' and ds_fpcode in " & codelist
    paydb.Execute sql
    
    sql = "update bio_device_shiftlogs set ds_status = 'WOHP' from bio_device_shiftlogs , bio_empmas  WHERE ds_fpcode = bioemp_fpcode and ds_sft_hrs >= 7.30 and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and  ds_status = 'WOH'  and ds_shift_actual not in ('A SHIFT','B SHIFT','C SHIFT')   and ds_fpcode in " & codelist
    paydb.Execute sql
    sql = "update bio_device_shiftlogs set ds_status = 'P' from bio_device_shiftlogs , bio_empmas WHERE ds_fpcode = bioemp_fpcode and bioemp_team = 'STAFF' and ds_sft_hrs >= 7.30 and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "'  and  ds_shift <> 'H' and  ds_shift <> 'WOH'  and  ds_shift <> 'WO'  and ds_shift_actual not in ('A SHIFT','B SHIFT','C SHIFT')  and bioemp_gender = 'F' and ds_fpcode in " & codelist
    paydb.Execute sql
    
    
    '' for  STAFF - MALE

''FOR STAFF -  PRESENT - for SINGLE IN / OUT punches

    sql = "update bio_device_shiftlogs set ds_status = 'P' from bio_device_shiftlogs , bio_empmas  WHERE ds_fpcode = bioemp_fpcode and bioemp_team = 'STAFF' and ds_sft_hrs >= 8.30 and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "'  and  ds_shift <> 'H'  and  ds_shift <> 'WOH' and ds_no_of_punches = 1 and ds_shift_actual not in ('A SHIFT','B SHIFT','C SHIFT')  and bioemp_gender = 'M' and ds_fpcode in " & codelist
    paydb.Execute sql
    sql = "update bio_device_shiftlogs set ds_status = 'WOP' from bio_device_shiftlogs , bio_empmas  WHERE ds_fpcode = bioemp_fpcode and bioemp_team = 'STAFF' and  ds_sft_hrs >= 8.30 and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and  ds_shift = 'WO' and ds_status = 'P' and ds_no_of_punches = 1   and ds_shift_actual not in ('A SHIFT','B SHIFT','C SHIFT')  and bioemp_gender = 'M' and ds_fpcode in " & codelist
    paydb.Execute sql
    
    sql = "update bio_device_shiftlogs set ds_status = 'WOHP' from bio_device_shiftlogs , bio_empmas  WHERE ds_fpcode = bioemp_fpcode and ds_sft_hrs >= 7.30 and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and  ds_status = 'WOH'  and ds_no_of_punches = 1 and ds_shift_actual not in ('A SHIFT','B SHIFT','C SHIFT')  and bioemp_gender = 'M' and ds_fpcode in " & codelist
    paydb.Execute sql
    sql = "update bio_device_shiftlogs set ds_status = 'P' from bio_device_shiftlogs , bio_empmas WHERE ds_fpcode = bioemp_fpcode and bioemp_team = 'STAFF' and ds_sft_hrs >= 8.30 and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "'  and  ds_shift <> 'H' and  ds_shift <> 'WOH'  and  ds_shift <> 'WO' and ds_no_of_punches = 1  and ds_shift_actual not in ('A SHIFT','B SHIFT','C SHIFT')  and bioemp_gender = 'M' and ds_fpcode in " & codelist
    paydb.Execute sql
    
    sql = "update bio_device_shiftlogs set ds_status = 'P' from bio_device_shiftlogs , bio_empmas  WHERE ds_fpcode = bioemp_fpcode and bioemp_team = 'STAFF' and ds_sft_hrs >= 8.30 and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "'  and  ds_shift <> 'H'  and  ds_shift <> 'WOH' and ds_shift_actual not in ('A SHIFT','B SHIFT','C SHIFT')  and bioemp_gender = 'M' and ds_fpcode in " & codelist
    paydb.Execute sql
    sql = "update bio_device_shiftlogs set ds_status = 'WOP' from bio_device_shiftlogs , bio_empmas  WHERE ds_fpcode = bioemp_fpcode and bioemp_team = 'STAFF' and  ds_sft_hrs >= 8.30 and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and  ds_shift = 'WO' and ds_status = 'P'   and ds_shift_actual not in ('A SHIFT','B SHIFT','C SHIFT')  and bioemp_gender = 'M' and ds_fpcode in " & codelist
    paydb.Execute sql
    
    sql = "update bio_device_shiftlogs set ds_status = 'WOHP' from bio_device_shiftlogs , bio_empmas  WHERE ds_fpcode = bioemp_fpcode and ds_sft_hrs >= 7.30 and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and  ds_status = 'WOH'  and ds_shift_actual not in ('A SHIFT','B SHIFT','C SHIFT')  and bioemp_gender = 'M' and ds_fpcode in " & codelist
    paydb.Execute sql
    sql = "update bio_device_shiftlogs set ds_status = 'P' from bio_device_shiftlogs , bio_empmas WHERE ds_fpcode = bioemp_fpcode and bioemp_team = 'STAFF' and ds_sft_hrs >= 8.30 and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "'  and  ds_shift <> 'H' and  ds_shift <> 'WOH'  and  ds_shift <> 'WO'  and ds_shift_actual not in ('A SHIFT','B SHIFT','C SHIFT')  and bioemp_gender = 'M' and ds_fpcode in " & codelist
    paydb.Execute sql
    
''FOR STAFF -  PRESENT - for DOUBLE IN / OUT punches

    sql = "update bio_device_shiftlogs set ds_status = 'P' from bio_device_shiftlogs , bio_empmas  WHERE ds_fpcode = bioemp_fpcode and bioemp_team = 'STAFF' and ds_sft_hrs >= 7.59 and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "'  and  ds_shift <> 'H'  and  ds_shift <> 'WOH' and ds_no_of_punches = 2 and ds_shift_actual not in ('A SHIFT','B SHIFT','C SHIFT')  and bioemp_gender = 'M' and ds_fpcode in " & codelist
    paydb.Execute sql
    sql = "update bio_device_shiftlogs set ds_status = 'WOP' from bio_device_shiftlogs , bio_empmas  WHERE ds_fpcode = bioemp_fpcode and bioemp_team = 'STAFF' and  ds_sft_hrs >= 7.59 and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and  ds_shift = 'WO' and ds_status = 'P' and ds_no_of_punches = 2   and ds_shift_actual not in ('A SHIFT','B SHIFT','C SHIFT')  and bioemp_gender = 'M' and ds_fpcode in " & codelist
    paydb.Execute sql
    
    sql = "update bio_device_shiftlogs set ds_status = 'WOHP' from bio_device_shiftlogs , bio_empmas  WHERE ds_fpcode = bioemp_fpcode and ds_sft_hrs >= 7.59 and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and  ds_status = 'WOH'  and ds_no_of_punches = 2 and ds_shift_actual not in ('A SHIFT','B SHIFT','C SHIFT')  and bioemp_gender = 'M' and ds_fpcode in " & codelist
    paydb.Execute sql
    
    
    
    
    
    
    
    sql = "update bio_device_shiftlogs set ds_status = 'P' from bio_device_shiftlogs , bio_empmas  WHERE ds_fpcode = bioemp_fpcode and bioemp_team <> 'STAFF' and ds_sft_hrs >= 7.45 and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "'  and  ds_shift <> 'H'  and  ds_shift <> 'WOH' and ds_shift_actual not in ('A SHIFT','B SHIFT','C SHIFT') and ds_fpcode in " & codelist
    paydb.Execute sql
    sql = "update bio_device_shiftlogs set ds_status = 'WOP' from bio_device_shiftlogs , bio_empmas  WHERE ds_fpcode = bioemp_fpcode and bioemp_team <> 'STAFF' and  ds_sft_hrs >= 7.45 and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and  ds_shift = 'WO' and ds_status = 'P' and ds_shift_actual not in ('A SHIFT','B SHIFT','C SHIFT') and ds_fpcode in " & codelist
    paydb.Execute sql
    sql = "update bio_device_shiftlogs set ds_status = 'WOHP' from bio_device_shiftlogs , bio_empmas  WHERE ds_fpcode = bioemp_fpcode and bioemp_team <> 'STAFF' and ds_sft_hrs >= 7.45 and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and  ds_status = 'WOH' and ds_shift_actual not in ('A SHIFT','B SHIFT','C SHIFT') and ds_fpcode in " & codelist
    paydb.Execute sql
    sql = "update bio_device_shiftlogs set ds_status = 'P' from bio_device_shiftlogs , bio_empmas WHERE ds_fpcode = bioemp_fpcode and bioemp_team <> 'STAFF' and ds_sft_hrs >= 7.45 and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "'  and  ds_shift <> 'H' and  ds_shift <> 'WOH'  and  ds_shift <> 'WO' and ds_shift_actual not in ('A SHIFT','B SHIFT','C SHIFT') and ds_fpcode in " & codelist
    paydb.Execute sql
    
    
    
    
    
''''FOR SECURITY GUARD
''    sql = "update bio_device_shiftlogs set ds_status = '½P'  from bio_device_shiftlogs , bio_empmas   WHERE  ds_fpcode = bioemp_fpcode and bioemp_team = 'SECURITY GUARD' and   ds_sft_hrs >4  and  ds_sft_hrs < 11.30 and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "'"
''    paydb.Execute sql
''    sql = "update bio_device_shiftlogs set ds_status = 'P'  from bio_device_shiftlogs , bio_empmas   WHERE  ds_fpcode = bioemp_fpcode and bioemp_team = 'SECURITY GUARD' and   ds_sft_hrs >=11.30  and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "'"
''    paydb.Execute sql
    
'' for OD Assignment
    sql = "update bio_device_shiftlogs set ds_status = 'P(OD)' from bio_device_shiftlogs , bio_empmas  WHERE ds_fpcode = bioemp_fpcode and  ds_sft_hrs >= 7.49 and  ds_od_hrs >= 4   and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and ds_fpcode in " & codelist
    paydb.Execute sql
    sql = "update bio_device_shiftlogs set ds_status = 'WOP(OD)' from bio_device_shiftlogs , bio_empmas  WHERE ds_fpcode = bioemp_fpcode and ds_sft_hrs >= 7.49 and  ds_od_hrs >= 4 and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and  ds_shift = 'WO' and ds_fpcode in " & codelist
    paydb.Execute sql
    
    
    sql = "update bio_device_shiftlogs set ds_status = 'WOHP' from bio_device_shiftlogs , bio_empmas  WHERE ds_fpcode = bioemp_fpcode and ds_sft_hrs >= 7.45 and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and  ds_status = 'WOH' and ds_fpcode in " & codelist
    paydb.Execute sql
    
    
    
    
    Dim leave, leavetype As String
 ''for updating leave entries
    pst_qry = "select * from bio_empleave where emp_leave_date between '" & Format(st_date, "MM/dd/yyyy") & "' and '" & Format(end_date, "MM/dd/yyyy") & "'"
    pst_qry = "select * from bio_empleave , bio_device_shiftlogs  where ds_fpcode = emp_fpcode and ds_date= emp_leave_date and emp_leave_date between '" & Format(st_date, "MM/dd/yyyy") & "' and '" & Format(end_date, "MM/dd/yyyy") & "' and ds_fpcode in " & codelist & " order by emp_leave_date"

    payrs.Open pst_qry, paydb, adOpenDynamic, adLockOptimistic
    While Not payrs.EOF
          id = payrs!emp_fpcode
''          If id = 1211 Then
''            If payrs("ds_date") = "29-01-2025" Then
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
               ElseIf payrs!emp_leave_type = "L" Then
                  leavetype = IIf(payrs!emp_leave_period = "F", "L", "½L")
               ElseIf payrs!emp_leave_type = "ML" Then
                  leavetype = IIf(payrs!emp_leave_period = "F", "ML", "½ML")
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
          ElseIf payrs!ds_status = "P" And leavetype <> "SA" Then
             sql = "update bio_device_shiftlogs set ds_status = 'P' where ds_fpcode =  '" & id & "' and ds_date =  '" & Format(payrs!emp_leave_date, "MM/dd/yyyy") & "'"
          ElseIf payrs!ds_status = "P" And leavetype = "SA" Then
             sql = "update bio_device_shiftlogs set ds_status = 'SA' where ds_fpcode =  '" & id & "' and ds_date =  '" & Format(payrs!emp_leave_date, "MM/dd/yyyy") & "'"
             
          ElseIf payrs!ds_status <> "A" And payrs!ds_status <> "" Then
             sql = "update bio_device_shiftlogs set ds_status = ds_status+'" & leavetype & "' where ds_fpcode =  '" & id & "' and ds_date =  '" & Format(payrs!emp_leave_date, "MM/dd/yyyy") & "'"
          ElseIf payrs!ds_status = "A" And leavetype = "½EL" Then
             sql = "update bio_device_shiftlogs set ds_status = '½EL½A' where ds_fpcode =  '" & id & "' and ds_date =  '" & Format(payrs!emp_leave_date, "MM/dd/yyyy") & "'"
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
    pst_qry = "select * from bio_emp_chleave , bio_device_shiftlogs  where ds_fpcode = empch_fpcode and ds_date= empch_ch_date and empch_ch_date between '" & Format(st_date, "MM/dd/yyyy") & "' and '" & Format(end_date, "MM/dd/yyyy") & "' and ds_fpcode in " & codelist

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
  sql = "update bio_device_shiftlogs set ds_status = '½P½A' from bio_device_shiftlogs where ds_Status = '½P' and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and ds_fpcode in " & codelist
  paydb.Execute sql
 
 
'''''for updating holiday present eligibility
'''''Start
'''
'''
'''  sql = "update bio_device_shiftlogs set ds_status = 'HPE' from bio_device_shiftlogs , bio_empmas ,emp_mas where bioemp_fpcode = emp_fpcode and bioemp_fpcode = ds_fpcode and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and emp_cat = 'W' and ds_status = 'HP' and emp_dh_wages_yn = 'Y'  and emp_status = 'A' and ds_fpcode in " & codelist
'''  paydb.Execute sql
'''
'''
'''  sql = "update bio_device_shiftlogs set ds_status = 'HPE' from bio_device_shiftlogs a, bio_empdh_eligible b where empdh_date = ds_date and  empdh_fpcode = ds_fpcode and ds_Status = 'HP' and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and ds_fpcode in " & codelist
'''  paydb.Execute sql
'''
'''
'''  sql = "update bio_device_shiftlogs set ds_status = 'HPE' from bio_device_shiftlogs , bio_empmas ,emp_mas where bioemp_fpcode = emp_fpcode and bioemp_fpcode = ds_fpcode and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and emp_cat = 'W' and ds_status = 'P(OD)'  and ds_shift = 'H' and emp_dh_wages_yn = 'Y' and emp_status = 'A' and ds_fpcode in " & codelist
'''  paydb.Execute sql
''''''MODIFIED BY DEVA'''''''
'''  sql = "update bio_device_shiftlogs set ds_status = 'HPE' from bio_device_shiftlogs a, bio_empdh_eligible b where empdh_date = ds_date and  empdh_fpcode = ds_fpcode and ds_Status = 'OD' and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and ds_fpcode in " & codelist
'''  paydb.Execute sql
'''  ''''''''''''
'''
'''  sql = "update bio_device_shiftlogs set ds_status = 'HPE' from bio_device_shiftlogs a, bio_empdh_eligible b where empdh_date = ds_date and  empdh_fpcode = ds_fpcode and ds_Status = 'P(OD)' and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and ds_fpcode in " & codelist
'''  paydb.Execute sql
'''
'''
'''  sql = "update bio_device_shiftlogs set ds_status = 'WOHPE' from bio_device_shiftlogs a, bio_empdh_eligible b where empdh_date = ds_date and  empdh_fpcode = ds_fpcode and ds_Status = 'WOHP' and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and ds_fpcode in " & codelist
'''  paydb.Execute sql
'''
'''  sql = "update bio_device_shiftlogs set ds_status = 'WOHPE' from bio_device_shiftlogs , bio_empmas ,emp_mas where bioemp_fpcode = emp_fpcode and bioemp_fpcode = ds_fpcode and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and emp_cat = 'W' and ds_status = 'WOHP'   and emp_status = 'A' and ds_fpcode in " & codelist
'''  paydb.Execute sql
'''
'''''End
 
''for 1/2CH and 1/2 ABS
  sql = "update bio_device_shiftlogs set ds_status = '½C.H½A' from bio_device_shiftlogs where ds_Status = '½C.H' and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and ds_sft_hrs = 0 and ds_fpcode in " & codelist
  paydb.Execute sql
 
 ''for 1/2CH and 1/2 PRESENT
  sql = "update bio_device_shiftlogs set ds_status = '½C.H½P' from bio_device_shiftlogs where ds_Status = '½C.H' and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and ds_sft_hrs > 3.4 and ds_fpcode in " & codelist
  paydb.Execute sql
 
 
 
'' FOR SATURDAY ATTENDANCE REMOVE

    pst_qry = "update bio_device_shiftlogs set ds_status = 'C' from bio_device_shiftlogs ,bio_emp_saturday where ds_fpcode = empsat_fpcode and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and '" & Format(end_date, "MM/dd/yyyy") & "' and datepart(dw,ds_date) =  7 and ds_status <> 'H'"
    pst_qry = "update bio_device_shiftlogs set ds_status = 'C' from bio_device_shiftlogs ,bio_emp_saturday where ds_fpcode = empsat_fpcode and ds_date =empsat_date and ds_fpcode in " & codelist
    
    paydb.Execute pst_qry
 

'' New Addditon on 28/10/2025 for more then 9.05 AM in punch
    sql = "update bio_device_shiftlogs set ds_status = '½P½A' from bio_device_shiftlogs , bio_empmas WHERE ds_fpcode = bioemp_fpcode and ds_sft_hrs >= 4.00 and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "'  and  ds_shift <> 'H' and  ds_shift <> 'WOH'  and  ds_shift <> 'WO'  and ds_shift_actual not in ('A SHIFT','B SHIFT','C SHIFT')   AND CONVERT(VARCHAR(8), ds_shift_in, 108) > '09:00:58' AND CONVERT(VARCHAR(8), ds_shift_in, 108) < '12:35:00' AND ds_date >= '2025-10-27' and ds_fpcode in " & codelist
    paydb.Execute sql
    
    sql = "update bio_device_shiftlogs set ds_status = 'P' from  bio_device_shiftlogs , bio_emp_oddetails  WHERE ds_fpcode = empod_fpcode  and empod_date = ds_date and   ds_date between '" & Format(st_date, "MM/dd/yyyy") & "'  and  '" & Format(end_date, "MM/dd/yyyy") & "' AND CONVERT(VARCHAR(8), ds_shift_in, 108) > '09:00:01'  AND CONVERT(VARCHAR(8), ds_shift_in, 108) < '12:35:00'  and  ds_shift <> 'H' and  ds_shift <> 'WOH'  and  ds_shift <> 'WO'  and ds_shift_actual not in ('A SHIFT','B SHIFT','C SHIFT') and ds_date >= '2025-10-27'  and empod_fromtime > 8 and empod_fromtime < 10 and ds_sft_hrs > 7 and ds_fpcode in " & codelist
    paydb.Execute sql

'' New Addditon on 17/11/2025 for A SHIFT
'' A SHIFT
    sql = "update bio_device_shiftlogs set ds_status = '½P½A' from bio_device_shiftlogs , bio_empmas WHERE ds_fpcode = bioemp_fpcode and ds_sft_hrs >= 4.00 and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "'  and  ds_shift <> 'H' and  ds_shift <> 'WOH'  and  ds_shift <> 'WO'  and ds_shift_actual not in ('A SHIFT')   AND CONVERT(VARCHAR(8), ds_shift_in, 108) > '06:00:58' AND CONVERT(VARCHAR(8), ds_shift_in, 108) < '07:00:00' AND ds_date >= '2025-11-18' and bioemp_dept <> 'DRIVER' and ds_fpcode in " & codelist
    paydb.Execute sql

'' B SHIFT
    sql = "update bio_device_shiftlogs set ds_status = '½P½A' from bio_device_shiftlogs , bio_empmas WHERE ds_fpcode = bioemp_fpcode and ds_sft_hrs >= 4.00 and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "'  and  ds_shift <> 'H' and  ds_shift <> 'WOH'  and  ds_shift <> 'WO'  and ds_shift_actual not in ('B SHIFT')   AND CONVERT(VARCHAR(8), ds_shift_in, 108) > '14:00:58' AND CONVERT(VARCHAR(8), ds_shift_in, 108) < '15:00:00' AND ds_date >= '2025-11-18'  and bioemp_dept <> 'DRIVER'  and ds_fpcode in " & codelist
    paydb.Execute sql
'' C SHIFT
    sql = "update bio_device_shiftlogs set ds_status = '½P½A' from bio_device_shiftlogs , bio_empmas WHERE ds_fpcode = bioemp_fpcode and ds_sft_hrs >= 4.00 and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "'  and  ds_shift <> 'H' and  ds_shift <> 'WOH'  and  ds_shift <> 'WO'  and ds_shift_actual not in ('C SHIFT')   AND CONVERT(VARCHAR(8), ds_shift_in, 108) > '22:00:58' AND CONVERT(VARCHAR(8), ds_shift_in, 108) < '23:30:00' AND ds_date >= '2025-11-18'  and bioemp_dept <> 'DRIVER' and ds_fpcode in " & codelist
    paydb.Execute sql
'' 6 PM TO 6AM
    sql = "update bio_device_shiftlogs set ds_status = '½P½A' from bio_device_shiftlogs , bio_empmas WHERE ds_fpcode = bioemp_fpcode and ds_sft_hrs >= 4.00 and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "'  and  ds_shift <> 'H' and  ds_shift <> 'WOH'  and  ds_shift <> 'WO'  and ds_shift_actual not in ('06.00PM-06.00AM')   AND CONVERT(VARCHAR(8), ds_shift_in, 108) > '18:00:58' AND CONVERT(VARCHAR(8), ds_shift_in, 108) < '19:30:00' AND ds_date >= '2025-11-18' and ds_fpcode in " & codelist
    paydb.Execute sql
 
 
 
 
''for updating bio_attendlogs_daily
    Dim aday As String
    sql = "select * from bio_device_shiftlogs where ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and ds_fpcode in " & codelist & " order by ds_date"
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
    Dim present, absent, hop, wop, cl, sl, h, ch, layoff, wo, pl, hope, el, woh, ml, emer_leave, wohp   As Single
    
    
    Dim blankdata, leavedata As Integer
    
    sql = "select * from bio_attendlogs where a_year = " & Val(cmb_year.Text) & "  and a_month = " & cmb_month.ItemData(cmb_month.ListIndex) & "  and a_fpcode in " & codelist
    payrs.Open sql, paydb, 1, 2
    If Not payrs.EOF Then
       While Not payrs.EOF
''            If payrs.Fields("a_fpcode") = 1489 Then
''               MsgBox ("Wait")
''            End If
            absent = 0
            blankdata = 0
            leavedata = 0
            For i = 1 To 31
            
''            If i = 15 Then
''                MsgBox ("Wait")
''            End If
            
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
''''                       MsgBox (payrs.Fields("a_fpcode"))
''                       If payrs.Fields(dayfind) = "WO" Then
''                           payrs.Fields(dayfind) = "A"
''                           payrs.Update
''                       End If
''                    End If
                    
                    If absent >= 3 And payrs.Fields(dayfind) = "WO" Then
                           payrs.Fields(dayfind) = "A"
''                            absent = 0
                           payrs.Update
                    End If
''                    If leavedata >= 3 And payrs.Fields(dayfind) = "WO" Then
''                           payrs.Fields(dayfind) = ""
''                            absent = 0
''                           payrs.Update
''                    End If
''                    If (blankdata + absent + leavedata) >= 6 And payrs.Fields(dayfind) = "WO" Then
                    If (blankdata + absent) >= 3 And payrs.Fields(dayfind) = "WO" Then
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
                           ''absent = 0
                    End If
                    If (blankdata + absent) >= 5 And payrs.Fields(dayfind) = "H" Then
                           payrs.Fields(dayfind) = ""
                           blankdata = 0
''                           absent = 0
                           payrs.Update
                    End If
                    If (blankdata + absent + leavedata) >= 3 And payrs.Fields(dayfind) = "WO" Then
                           payrs.Fields(dayfind) = ""
                           blankdata = 0
''                           absent = 0
                           leavedata = 0
                           payrs.Update
                    End If
''End
                    
                    
                End If
            Next
            payrs.MoveNext
        Wend
    End If
    payrs.Close
    
    
 
    sql = "select * from bio_attendlogs  ,emp_mas where a_fpcode = emp_fpcode and a_year = " & Val(cmb_year.Text) & "  and a_month = " & cmb_month.ItemData(cmb_month.ListIndex) & "  and a_fpcode in " & codelist
    
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
            wohp = 0
            ml = 0
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
                ElseIf payrs.Fields(dayfind) = "WO" Or payrs.Fields(dayfind) = "WOL" Then
                    wo = wo + 1
                ElseIf payrs.Fields(dayfind) = "WO½P" Then
                    wop = wop + 0.5
                    wo = wo + 0.5
                ElseIf payrs.Fields(dayfind) = "WOH" Then
                    woh = woh + 1
                    wo = wo + 1
 ''                   h = h + 1
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
                    
                ElseIf payrs.Fields(dayfind) = "½PC.H" Then
                    ch = ch + 0.5
                    present = present + 0.5
                ElseIf payrs.Fields(dayfind) = "½PC.H½C.H" Then
                    ch = ch + 0.5
                    present = present + 0.5
                ElseIf payrs.Fields(dayfind) = "P½C.H" Then
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
            payrs("a_wo") = wo + woh
            payrs("a_woh") = woh
            payrs("a_emer_leave_days") = emer_leave
            payrs("a_month_days") = mdays
            
''            If payrs("emp_cat") = "W" Then
''               totdays = present + wop + wo + wop + ch + ml
''            Else
''               totdays = present + wop + wo + ch + ml
''            End If
            
      ''      MsgBox (wop)
            
            If payrs("emp_cat") = "W" Then
               totdays = present + wop + wo + wop + ch + ml + hop
            Else
               totdays = present + wop + wo + ch + ml + woh
            End If
                                  
            
'''            payrs("a_month_days") = mdays
'''
'''''            totdays = present + wop + wo + wop + ch
'''
'''            payrs("a_total_days") = totdays
'''
'''
'''
'''            If totdays >= mdays Then
'''                payrs("a_eligible_days") = mdays
'''                payrs("a_salary_days") = mdays + h
'''                If payrs("emp_cat") = "W" Then
'''                   payrs("a_ot_days") = totdays - mdays
'''                Else
'''                   payrs("a_ot_days") = 0
'''                End If
'''           Else
'''
'''                payrs("a_eligible_days") = totdays
'''                payrs("a_salary_days") = totdays + h
'''                payrs("a_ot_days") = 0
'''            End If
            
            payrs("a_month_days") = mdays
            
''            totdays = present + wop + wo + wop + ch
            
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
    sql = "select * ,case when emp_ch_period = 'F' then 1 else 0.5 end as Holidaych from bio_device_shiftlogs a, bio_emp_chleave b where  ds_fpcode =  empch_fpcode and ds_shift = 'H' and ds_date = empch_worked_date and ds_date Between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and ds_fpcode in " & codelist & " order by ds_date"
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
 
 
   
   
   
   
   
   MousePointer = vbDefault
   gst_repconnect = "dsn=pay_new;uid=sa;pwd=serdat"
''   cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\payroll\payslip.rpt"
''   cry_rep1.Formulas(0) = ("report_month = " & cmb_month.Text)
   cry_rep1.Formulas(0) = ("report_month = '" & cmb_month.Text & "'")
   cry_rep1.Formulas(1) = ("report_year = '" & cmb_year.Text & "'")
   cry_rep1.Formulas(2) = ("millname= '" & millname & "'")
   cry_rep1.PrinterSelect
   
   
   If opt_vou.Value = True Then
      cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\payroll\monthly_attendance_status_vou.rpt"
      cry_rep1.ReplaceSelectionFormula ("{emp_voupay_mast.emp_status} = 'A' and {bio_attendlogs.a_year} = " & Val(cmb_year.Text) & " and {bio_attendlogs.a_month}= " & cmb_month.ItemData(cmb_month.ListIndex) & " and (" & sel_codes & ")")
   ElseIf opt_regular.Value = True Then
      cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\payroll\monthly_attendance_status_deptwise.rpt"
      cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\payroll\monthly_attendance_status_deptwise_full.rpt"
      cry_rep1.ReplaceSelectionFormula ("{emp_mas.emp_status} = 'A' and {bio_attendlogs.a_year} = " & Val(cmb_year.Text) & " and {bio_attendlogs.a_month}= " & cmb_month.ItemData(cmb_month.ListIndex) & " and (" & sel_codes & ")")
    
   Else
     cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\cs\monthly_attendance_status.rpt"
      cry_rep1.ReplaceSelectionFormula ("{mas_caemp.ca_status} = 'A' and {bio_attendlogs.a_year} = " & Val(cmb_year.Text) & " and {bio_attendlogs.a_month}= " & cmb_month.ItemData(cmb_month.ListIndex) & " and (" & sel_codes & ")")

   End If
   
   cry_rep1.WindowState = crptMaximized
   cry_rep1.Connect = gst_repconnect
   cry_rep1.Action = 1
    
    
   Exit Sub


err_handler:
     chk = gen_Validation(Err.Number, Err.Description)
    '' paydb.RollbackTrans
     Me.MousePointer = 1
  '  chk = gen_Validation(Err.Number, Err.Description)
      If chk = 1 Then Resume
    

End Sub


Private Sub Command1_Click()
     Dim t1_1, t2_1, t1_2, t2_2, t1, t2 As Double
     pst_qry = "select * from bio_emp_oddetails where empod_date between '" & Format(st_date, "MM/dd/yyyy") & "' and '" & Format(end_date, "MM/dd/yyyy") & "'"
     pst_qry = "select * from bio_emp_oddetails "
    payrs.Open pst_qry, paydb, adOpenDynamic, adLockOptimistic
    While Not payrs.EOF
          id = payrs!empod_fpcode
          idate = payrs!empod_date
          t1_1 = Left(payrs!empod_fromtime, 2)
          t1_2 = Mid(payrs!empod_fromtime, 4, 2)
          t2_1 = Left(payrs!empod_totime, 2)
          t2_2 = Mid(payrs!empod_totime, 4, 2)
          t1 = t1_1 + IIf(t1_2 > 0, t1_2 / 100, 0)
          t2 = t2_1 + IIf(t2_2 > 0, t2_2 / 100, 0)
                payrs("empod_fromtime2") = t1
                payrs("empod_totime2") = t2
                payrs.Update
                payrs.MoveNext
     
     Wend
     MsgBox ("Updated..")
     payrs.Close
End Sub

Private Sub exit_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    del_leave = 0

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
''''        .AddItem finyear + 2000
''        .AddItem "2015"
''        .AddItem "2016"
''        .Text = "2015"
      .AddItem Left(fyear, 4)
      .AddItem Mid(fyear, 6, 4)
      If Year(Date) = Int(Left(fyear, 4)) Then
         cmb_year.Text = Left(fyear, 4)
      Else
          cmb_year.Text = Mid(fyear, 6, 4)
      End If

    End With

    cmb_month.ListIndex = Month(Date) - 1
    
    Dim payrs As New ADODB.Recordset
    lst_dept.Clear

    sql = "select bioemp_dept  from bio_empmas where bioemp_status = 'Working' group by bioemp_dept order by bioemp_dept"
    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
    While Not payrs.EOF()
        lst_dept.AddItem payrs("bioemp_dept")
        payrs.MoveNext
    Wend
    payrs.Close

    ''Dim itmX As ListItem
    Dim itmX As MSComctlLib.ListItem
    lst_view.ColumnHeaders.Clear
    lst_view.ColumnHeaders.Add , , "FP Code ", 1000
    lst_view.ColumnHeaders.Add , , "Emp. Name ", 2000
    lst_view.ColumnHeaders.Add , , "Department ", 1500
    lst_view.ColumnHeaders.Add , , "Type ", 1000
    lst_view.View = lvwReport
    lst_view.ListItems.Clear
    
    sql = "select * from bio_empmas where bioemp_status = 'Working' order by bioemp_dept"
    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
    While Not payrs.EOF()
            Set itmX = lst_view.ListItems.Add(, , CStr(payrs("bioemp_fpcode")))
            itmX.SubItems(1) = payrs.Fields("bioemp_name")
            itmX.SubItems(2) = payrs.Fields("bioemp_dept")
            itmX.SubItems(3) = payrs.Fields("bioemp_team")
            payrs.MoveNext
    Wend
    
    payrs.Close
    
find_dates
End Sub
Private Sub cmb_month_Click()
    find_dates
End Sub

Private Sub cmb_year_Click()
   find_dates
    
End Sub
Public Sub find_dates()
''    If cmb_month.ListIndex = -1 Then Exit Sub
''    Dim d1 As Date
''    mmon = cmb_month.ItemData(cmb_month.ListIndex)
''    If mmon = 1 Or mmon = 3 Or mmon = 5 Or mmon = 7 Or mmon = 8 Or mmon = 10 Or mmon = 12 Then
''        mdays = 31
''    ElseIf mmon = 4 Or mmon = 6 Or mmon = 9 Or mmon = 11 Then
''        mdays = 30
''    ElseIf mmon = 2 And Val(cmb_year.Text) Mod 4 = 0 Then
''        mdays = 29
''    Else
''        mdays = 28
''    End If
''    end_date = DateValue(Str(mmon) + "/" + Str(mdays) + "/" + cmb_year.Text) + 1
''    st_date = end_date - Day(end_date) + 1
''    If end_date.Value > Now Then
''       end_date.Value = Now + 1
''    End If

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
    
    month_start_date.Value = st_date

    If end_date.Value > Now Then
       end_date.Value = Now + 1
    End If
End Sub

Private Sub lst_dept_Click()
    lst_employee.Clear
    sql = "Select * from  bio_empmas where bioemp_status = 'Working'  and bioemp_dept = '" & lst_dept.Text & "' order by bioemp_name"
    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
    While Not payrs.EOF()
        lst_employee.AddItem payrs("bioemp_name")
        lst_employee.ItemData(lst_employee.NewIndex) = payrs("bioemp_fpcode")
        payrs.MoveNext
    Wend
    payrs.Close
End Sub

Private Sub Refresh_Click()
    lst_view.Refresh
End Sub

