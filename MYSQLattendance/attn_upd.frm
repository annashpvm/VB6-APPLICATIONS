VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form attn_upd 
   Caption         =   "ATTENDANCE UPDATION S"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.CommandButton EXIT 
      Caption         =   "EXIT"
      Height          =   495
      Left            =   6720
      TabIndex        =   6
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton cmd_process 
      Caption         =   "UPDATION"
      Height          =   495
      Left            =   4800
      TabIndex        =   5
      Top             =   3000
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker end_date 
      Height          =   375
      Left            =   7320
      TabIndex        =   4
      Top             =   1440
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      Format          =   20643841
      CurrentDate     =   39359
   End
   Begin MSComCtl2.DTPicker st_date 
      Height          =   375
      Left            =   3360
      TabIndex        =   3
      Top             =   1440
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      Format          =   20643841
      CurrentDate     =   39359
   End
   Begin VB.Frame Frame1 
      Height          =   3855
      Left            =   1440
      TabIndex        =   0
      Top             =   1080
      Width           =   8775
      Begin VB.Label Label2 
         Caption         =   "End Date"
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
         Height          =   375
         Left            =   4320
         TabIndex        =   2
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Start Date"
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
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   1335
      End
   End
End
Attribute VB_Name = "attn_upd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public a_mill As String
Public a_type As String
Public a_date As Date
Public a_time As Date
Public a_id As String
Public a_uname As String
Public empcode As String

Public Sub trans_check()
    
    Dim mdbrs2 As New ADODB.Recordset
''        pstr_qry = "select * from history " _
''                   & "where eventdate >=# " _
''                   & Format(dtp_fromdate - 1, "dd/mm/yyyy") & "# and " _
''                   & "eventdate <=#" & Format(dtp_todate - 1, "dd/mm/yyyy") & "# and " _
''                   & "eventtime>=#23:00# and eventtime<=#23:59# and eventno = 4 order by eventdate,eventtime"
''
   mdb_qry2 = "select * from History where EVENTDATE >= #" & Format(st_date, "dd/mm/yyyy") & "# and EVENTDATE <= #" & Format(end_date, "dd/mm/yyyy") & "# and UserId = '" & empcode & "' order by eventdate,eventtime"
   mdbrs3.Open mdb_qry2, dsnmdb2, 1, 2
   While Not mdbrs2.EOF
              If Trim(mdbrs2(6)) <> "" Then
                 a_date = mdbrs2(0)
                 a_time = mdbrs2(1)
                 a_id = mdbrs2(5)
                 a_uname = mdbrs2(6)
                 dataupd
              End If
              mdbrs2.MoveNext
    Wend
End Sub


Public Sub dataupd()
    Dim attnrs As New ADODB.Recordset
    SQL_QRY = "select * from attendance"
    attnrs.Open SQL_QRY, attndb, adOpenDynamic, adLockOptimistic
    attnrs.AddNew
    attnrs("empmill") = a_mill
    attnrs("emptype") = a_type
    attnrs("empcode") = a_id
    attnrs("empname") = a_uname
    attnrs("attndate") = a_date
    attnrs("intime") = Format(a_time, "HH:MM:SS")
       attnrs("outtime") = Format(a_time, "HH:MM:SS")
       attnrs("timein") = a_time
       attnrs("timeout") = a_time
       attnrs("indate") = a_date
       attnrs("outdate") = a_date
    attnrs.UPDATE
End Sub



Private Sub cmd_process_Click()

    Dim attnrs As New ADODB.Recordset
    SQL_QRY = "delete from attendance where attndate >= '" & Format(st_date.Value, "mm-dd-yyyy") & "' and attndate <= '" & Format(end_date.Value, "mm-dd-yyyy") & "'"
    attnrs.Open SQL_QRY, attndb, adOpenDynamic, adLockOptimistic
    dsnmdb = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=\\personnel\plat04s\Acu301.mdb"
    Dim mdbrs As New ADODB.Recordset
            Set mdbrs3 = New ADODB.Recordset
'            mdb_qry2 = "select * from History where EVENTDATE >= #" & Format(st_date, "dd/mm/yyyy") & "# and EVENTDATE <= #" & Format(end_date, "dd/mm/yyyy") & "# and UserId = '" & empcode & "' order by eventdate,eventtime"
            mdb_qry2 = "select * from History where EVENTDATE >= #" & Format(st_date.Value, "mm/dd/yyyy") & "# and EVENTDATE <= #" & Format(end_date.Value, "mm/dd/yyyy") & "#  order by eventdate,eventtime"
            
            mdbrs3.Open mdb_qry2, dsnmdb, 1, 2
            While Not mdbrs3.EOF
                  If Trim(mdbrs3(6)) <> "" Then
                     a_type = ""
                     a_mill = Mid(mdbrs3(5), 4, 1)
                     a_date = Format(mdbrs3(0), "mm/dd/yyyy")
                     a_time = mdbrs3(1)
                     a_id = mdbrs3(5)
                     a_uname = mdbrs3(6)
                     Select Case Left(mdbrs3(5), 1)
                     Case 1
                          a_type = "S"
                     Case 5
                          a_type = "W"
                     End Select
                     
                     dataupd
                  End If
                  mdbrs3.MoveNext
            Wend
    MsgBox ("Records are updated FROM DPM ")
         dsnmdb = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=d:\temp\Acu301.mdb"

    ''dsnmdb = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=\\slpbto\plat04s\Acu301.mdb"
            Set mdbrs3 = New ADODB.Recordset
            mdb_qry2 = "select * from History where EVENTDATE >= #" & Format(st_date, "mm/dd/yyyy") & "# and EVENTDATE <= #" & Format(end_date, "mm/dd/yyyy") & "#  order by eventdate,eventtime"
            mdbrs3.Open mdb_qry2, dsnmdb, 1, 2
            While Not mdbrs3.EOF
                  If Trim(mdbrs3(6)) <> "" Then
                     a_type = ""
                     a_mill = Mid(mdbrs3(5), 4, 1)
                     a_date = Format(mdbrs3(0), "mm/dd/yyyy")
                     a_time = mdbrs3(1)
                     a_id = mdbrs3(5)
                     a_uname = mdbrs3(6)
                     Select Case Left(mdbrs3(5), 1)
                     Case 1
                          a_type = "S"
                     Case 5
                          a_type = "W"
                     End Select
                     
                     dataupd
                  End If
                  mdbrs3.MoveNext
            Wend
'    trans_check
    MsgBox ("Records are updated FROM SLPB ")
    dsnmdb = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=\\cogento\plat04s\Acu301.mdb"
    Set mdbrs = New ADODB.Recordset
''    Dim mdbrs As New ADODB.Recordset
''    mdb_qry = "Select * from Member"
''    mdbrs.Open mdb_qry, dsnmdb, adOpenDynamic, adLockOptimistic
''    totemp = 0
''    While Not mdbrs.EOF
''        If Left(mdbrs(0), 3) <> "999" Then
''            empcode = mdbrs(0)
            Set mdbrs3 = New ADODB.Recordset
'            mdb_qry2 = "select * from History where EVENTDATE >= #" & Format(st_date, "dd/mm/yyyy") & "# and EVENTDATE <= #" & Format(end_date, "dd/mm/yyyy") & "# and UserId = '" & empcode & "' order by eventdate,eventtime"
            mdb_qry2 = "select * from History where EVENTDATE >= #" & Format(st_date, "mm/dd/yyyy") & "# and EVENTDATE <= #" & Format(end_date, "mm/dd/yyyy") & "#  order by eventdate,eventtime"
            mdbrs3.Open mdb_qry2, dsnmdb, 1, 2
            While Not mdbrs3.EOF
                  If Trim(mdbrs3(6)) <> "" Then
                     a_type = ""
                     a_mill = Mid(mdbrs3(5), 4, 1)
                     a_date = Format(mdbrs3(0), "mm/dd/yyyy")
                     a_time = mdbrs3(1)
                     a_id = mdbrs3(5)
                     a_uname = mdbrs3(6)
                     Select Case Left(mdbrs3(5), 1)
                     Case 1
                          a_type = "S"
                     Case 5
                          a_type = "W"
                     End Select
                     dataupd
                  End If
                  mdbrs3.MoveNext
            Wend
''        End If
''        mdbrs.MoveNext
''    Wend
'    trans_check
    MsgBox ("Records are updated FROM COGEN ")
End Sub
Private Sub EXIT_Click()
       Unload Me
End Sub

Private Sub Form_Load()
'' dsnmdb = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=\\personnel\plat04s\Acu301.mdb"
 ''dsnmdb2 = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=\\personnel\plat04s\Acu301.mdb"
    st_date = Date
    end_date = Date
End Sub
