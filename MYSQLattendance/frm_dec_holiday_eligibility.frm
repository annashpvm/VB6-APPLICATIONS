VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frm_dec_holiday_eligibility 
   Caption         =   "Declare Holiday Eligibility"
   ClientHeight    =   8955
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15555
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8955
   ScaleWidth      =   15555
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Caption         =   "Select Department"
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
      Height          =   3975
      Left            =   960
      TabIndex        =   7
      Top             =   2040
      Width           =   11055
      Begin ComctlLib.ListView lst_view 
         Height          =   2655
         Left            =   3360
         TabIndex        =   13
         Top             =   480
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   4683
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   327682
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.CommandButton cmd_deselect 
         Caption         =   "Deselect All"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9000
         TabIndex        =   12
         Top             =   840
         Width           =   1695
      End
      Begin VB.CommandButton cmd_allselect 
         Caption         =   "Select All"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9000
         TabIndex        =   11
         Top             =   360
         Width           =   1695
      End
      Begin VB.CommandButton cmd_Assign 
         Caption         =   "Assign Eligibility"
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
         Left            =   9000
         TabIndex        =   10
         Top             =   2880
         Width           =   1695
      End
      Begin VB.CommandButton cmd_filter 
         Caption         =   "FILTER"
         Height          =   375
         Left            =   480
         TabIndex        =   9
         Top             =   3240
         Width           =   1215
      End
      Begin VB.ListBox lst_dept 
         Height          =   2595
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Width           =   2655
      End
   End
   Begin VB.Frame Frame5 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   4680
      TabIndex        =   4
      Top             =   6840
      Width           =   2175
      Begin VB.CommandButton exit 
         BackColor       =   &H00C0E0FF&
         Caption         =   "&Exit"
         Height          =   705
         Left            =   1080
         MaskColor       =   &H000000FF&
         Picture         =   "frm_dec_holiday_eligibility.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton Refresh 
         BackColor       =   &H00C0E0FF&
         Caption         =   "&Refresh"
         Height          =   705
         Left            =   120
         MaskColor       =   &H000000FF&
         Picture         =   "frm_dec_holiday_eligibility.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   120
         Width           =   735
      End
   End
   Begin VB.Frame Frame7 
      Height          =   855
      Left            =   2880
      TabIndex        =   0
      Top             =   960
      Width           =   6015
      Begin VB.ComboBox cmb_holiday 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   2640
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   240
         Width           =   2535
      End
      Begin VB.Label Label10 
         Caption         =   "Select Declare Holiday"
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
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   2175
      End
   End
   Begin VB.Label lbl_emp 
      Alignment       =   2  'Center
      Caption         =   "EMPLOYEE DECLARE HOLIDAY ELIGIBILITY"
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
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   10695
   End
End
Attribute VB_Name = "frm_dec_holiday_eligibility"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''Private Sub cmd_allselect_Click()
''    For i = 1 To lst_view.ListItems.Count
''        ''lst_view.ListItems(i).Checked = True
''    Next
''
''End Sub
''
''Private Sub cmd_Assign_Click()
''    If cmb_holiday.Text = "" Then
''       MsgBox ("Select Declare holiday date...")
''       Exit Sub
''    End If
''
''    Dim pst_qry As String
''    Dim payrs As New ADODB.Recordset
''
''
''''    Dim iSelected As Integer
''''    Dim item As ListItem
''''    For i = 1 To lst_view.ListItems.Count
''''        If lst_view.ListItems(i).Checked = True Then
''''          iSelected = iSelected + 1
''''        End If
''''    Next
''''    If iSelected = 0 Then
''''       MsgBox ("Employee Not selected in the view...")
''''       Exit Sub
''''    End If
''
''
''    paydb.BeginTrans
''
''    sql = " delete from  bio_empdh_eligible where empdh_date = '" & Format(cmb_holiday.Text, "MM/dd/yyyy") & "' and empdh_fpcode in (select bioemp_fpcode from bio_empmas  where bioemp_dept = '" & lst_dept.Text & "')"
''    paydb.Execute sql
''    For i = 1 To lst_view.ListItems.Count
''        If lst_view.ListItems(i).Checked = True Then
''           sql = "insert into bio_empdh_eligible (empdh_fpcode, empdh_date) values ( " & lst_view.ListItems(i).Text & ",'" & Format(cmb_holiday.Text, "MM/dd/yyyy") & "')"
''           paydb.Execute sql
''        End If
''    Next
''    paydb.CommitTrans
''    MsgBox "Record Saved", vbOKOnly + vbInformation, "Information"
''
''    Exit Sub
''Exit Sub
''err_handler:
''        paydb.RollbackTrans
''        chk = gen_Validation(Err.Number, Err.Description)
''
''
''End Sub
''
''Private Sub cmd_deselect_Click()
''    For i = 1 To lst_view.ListItems.Count
''        lst_view.ListItems(i).Checked = False
''    Next
''
''End Sub
''
''Private Sub cmd_filter_Click()
''    If cmb_holiday.Text = "" Then
''       MsgBox ("Select Declare holiday date...")
''       Exit Sub
''    End If
''    Refresh_Click
''    Dim payrs As New ADODB.Recordset
''    Dim itmx As ListItem
''    lst_view.ColumnHeaders.Clear
''    lst_view.ColumnHeaders.Add , , "FP Code ", 1000
''    lst_view.ColumnHeaders.Add , , "Emp. Name ", 2000
''    lst_view.ColumnHeaders.Add , , "Department ", 1500
''    lst_view.View = lvwReport
''    lst_view.ListItems.Clear
'' ''   sql = "select * from bio_empmas where bioemp_status = 'Working' and bioemp_dept = '" & lst_dept.Text & "' order by bioemp_dept"
''''    sql = "select  * from bio_empmas a, emp_mas b where bioemp_fpcode = emp_fpcode and emp_classification = 'B' and bioemp_status = 'Working' and bioemp_dept = '" & lst_dept.Text & "' and emp_cat = 'S'  and emp_status = 'A'  order by bioemp_dept"
''    sql = "select  bioemp_fpcode,bioemp_name,bioemp_dept,1 as checkedlist from bio_empmas a, emp_mas b,bio_empdh_eligible c where bioemp_fpcode = emp_fpcode and emp_fpcode=empdh_fpcode and convert(varchar(10),empdh_date,103)= '" & cmb_holiday.Text & "'  and emp_classification = 'B' and bioemp_status = 'Working' and bioemp_dept = '" & lst_dept.Text & "' and emp_cat = 'S'  and emp_status = 'A' " _
''           & " Union All " _
''          & " select  bioemp_fpcode,bioemp_name,bioemp_dept,0 as checkedlist   from bio_empmas a, emp_mas b where bioemp_fpcode = emp_fpcode and emp_classification = 'B' and bioemp_status = 'Working' and bioemp_dept = '" & lst_dept.Text & "' and emp_cat = 'S'  and emp_status = 'A' and bioemp_fpcode not in (select  bioemp_fpcode from bio_empmas a, emp_mas b,bio_empdh_eligible c where bioemp_fpcode = emp_fpcode and emp_fpcode=empdh_fpcode and convert(varchar(10),empdh_date,103)= '" & cmb_holiday.Text & "'  and emp_classification = 'B' and bioemp_status = 'Working' and bioemp_dept = '" & lst_dept.Text & "' and emp_cat = 'S'  and emp_status = 'A')  order by bioemp_dept"
''    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
''    While Not payrs.EOF()
''
''            Set itmx = lst_view.ListItems.Add(, , CStr(payrs("bioemp_fpcode")))
''            If payrs.Fields("checkedlist") = 1 Then
''                itmx.Checked = True
''            Else
''                itmx.Checked = False
''            End If
''            itmx.SubItems(1) = payrs.Fields("bioemp_name")
''            itmx.SubItems(2) = payrs.Fields("bioemp_dept")
''
''            payrs.MoveNext
''    Wend
''    payrs.Close
''End Sub
''
''Private Sub Command1_Click()
''
''End Sub
''
''Private Sub exit_Click()
''    Unload Me
''End Sub
''
''Private Sub Form_Load()
''    Dim payrs As New ADODB.Recordset
''
''
''    Dim fdate, edata As Date
''    fdate = DateValue("01/01/" + Str(Year(Now)))
''    edate = DateValue("12/31/" + Str(Year(Now)))
''''    dec_holiday = Date
''    Set paydb = New ADODB.Connection
''    Set payrs = New ADODB.Recordset
''    sql = "Select * from emp_dec_holiday where emp_dec_holiday between '" & Format(fdate, "MM/dd/yyyy") & "' and '" & Format(edate, "MM/dd/yyyy") & "'"
''    paydb.Open pay
''    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
''    While Not payrs.EOF
''        cmb_holiday.AddItem Format(payrs!emp_dec_holiday, "dd/MM/yyyy")
''        payrs.MoveNext
''    Wend
''    payrs.Close
''
''
''    lst_dept.Clear
''    sql = "select bioemp_dept  from bio_empmas where bioemp_status = 'Working' group by bioemp_dept order by bioemp_dept"
''    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
''    While Not payrs.EOF()
''        lst_dept.AddItem payrs("bioemp_dept")
''        payrs.MoveNext
''    Wend
''    payrs.Close
''
''
''End Sub
''
''Private Sub lst_dept_Click()
''     lst_view.ListItems.Clear
''End Sub
''
''Private Sub Refresh_Click()
''    lst_view.ListItems.Clear
''End Sub
''
''
