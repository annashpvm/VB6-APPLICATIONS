VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form attn_upd_new 
   Caption         =   "ATTENDANCE UPDATION "
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   9840
      Top             =   5400
   End
   Begin MSComCtl2.DTPicker rdate 
      Height          =   255
      Left            =   4680
      TabIndex        =   7
      Top             =   1320
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   450
      _Version        =   393216
      Format          =   16580609
      CurrentDate     =   39366
   End
   Begin VB.CommandButton exit 
      Caption         =   "EXIT"
      Height          =   975
      Left            =   4800
      TabIndex        =   3
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Select Date for Update"
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
      Height          =   3735
      Left            =   1080
      TabIndex        =   0
      Top             =   1560
      Width           =   7575
      Begin VB.CommandButton UPDATE 
         Caption         =   "UPDATE"
         Height          =   975
         Left            =   2160
         TabIndex        =   2
         Top             =   2040
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker end_date 
         Height          =   375
         Left            =   5640
         TabIndex        =   4
         Top             =   840
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   16580609
         CurrentDate     =   39359
      End
      Begin MSComCtl2.DTPicker st_date 
         Height          =   375
         Left            =   1800
         TabIndex        =   5
         Top             =   840
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   16580609
         CurrentDate     =   39359
      End
      Begin VB.Label Label2 
         Caption         =   "END DATE"
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
         Height          =   375
         Left            =   3960
         TabIndex        =   6
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "START DATE"
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
         Height          =   495
         Left            =   480
         TabIndex        =   1
         Top             =   840
         Width           =   1335
      End
   End
   Begin VB.Label msglabel 
      Caption         =   "DATA UPDATING FROM DPM UNIT"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   1200
      TabIndex        =   8
      Top             =   5760
      Width           =   6975
   End
End
Attribute VB_Name = "attn_upd_new"
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
Public inout As String
Dim cshift As String
Dim UID As String
Dim qry As String
Public SwitchVal As Boolean

Private Sub EXIT_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    st_date = Date
    end_date = Date
    rdate = Date
End Sub


Private Sub UPDATE_Click()
    msglabel.Caption = "DATA UPDATING FROM DPM UNIT"
    Set attnrs = New ADODB.Recordset
''    SQL_QRY = "delete from tmp_attn where attndate >= '" & Format(st_date.Value, "mm-dd-yyyy") & "' and attndate <= '" & Format(end_date.Value, "mm-dd-yyyy") & "'"
    sql_qry = "delete from attendance where attndate >= '" & Format(st_date.Value, "mm-dd-yyyy") & "' and attndate <= '" & Format(end_date.Value, "mm-dd-yyyy") & "'"
    
    attnrs.Open sql_qry, attndb, adOpenDynamic, adLockOptimistic
'coding for update from DPM / SLPB / COGEN Units
    For i = 1 To 3
        Select Case i
        Case 1
         dsnmdb = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=\\personnel\plat04s\Acu301.mdb"
''         dsnmdb = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=d:\Acu301.mdb"
         
''         dsnmdb = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=d:\to be deleted\Acu301.mdb"
        Case 2
''         dsnmdb = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=\\slpbto\plat04s\Acu301.mdb"
         dsnmdb = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=d:\temp\Acu301.mdb"
        Case 3
         dsnmdb = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=\\cogento\plat04s\Acu301.mdb"
        End Select
        Set mdbrs = New ADODB.Recordset
        qry = "select * from history where EVENTDATE >= #" & Format(st_date.Value - 1, "MM/dd/yyyy") & "# and EVENTDATE <= #" & Format(end_date.Value, "mm/dd/yyyy") & "# order by eventdate,eventtime"
        mdbrs.Open qry, dsnmdb, 1, 2
        While Not mdbrs.EOF
          If Trim(mdbrs(6)) <> "" Or Mid(mdbrs(6), 1, 1) = "0" Then
             a_type = ""
             a_mill = Mid(mdbrs(5), 4, 1)
             a_date = Format(mdbrs(0), "mm/dd/yyyy")
             a_time = mdbrs(1)
             a_id = mdbrs(5)
             a_uname = mdbrs(6)
             Select Case Left(mdbrs(5), 1)
             Case 1
                  a_type = "S"
             Case 5
                  a_type = "W"
             End Select
             inout = Trim(mdbrs("FuncCode"))
             dataupd
          End If
          mdbrs.MoveNext
        Wend
    Next
    
    MsgBox ("Records are updated")
End Sub


Public Sub dataupd()
       Dim attnrs As New ADODB.Recordset
''       a = Format(a_time, "HH:MM:SS")
''       MsgBox (a)
       out_date = a_date
       If inout = "10" And Format(a_time, "HH:MM:SS") > "04:00:00" And Format(a_time, "HH:MM:SS") < "10:00:00" Then
          a_date = a_date - 1
       End If
''       SQL_QRY = "select * from tmp_attn where empcode = " & a_id & " and attndate = '" & Format(a_date, "mm/dd/yyyy") & "'"
       sql_qry = "select * from attendance where empcode = " & a_id & " and attndate = '" & Format(a_date, "mm/dd/yyyy") & "'"
       attnrs.Open sql_qry, attndb, adOpenDynamic, adLockOptimistic
       If attnrs.EOF Then
          attnrs.AddNew
          attnrs("empmill") = a_mill
          attnrs("emptype") = a_type
          attnrs("empcode") = a_id
          attnrs("empname") = a_uname
          attnrs("attndate") = a_date
       '   If inout = "0" Then
             attnrs("intime") = Format(a_time, "HH:MM:SS")
             attnrs("timein") = a_time
             attnrs("indate") = a_date
'          End If
          'If inout = "10" Then
             attnrs("outtime") = Format(a_time, "HH:MM:SS")
             attnrs("timeout") = a_time
             attnrs("outdate") = out_date
'          End If
       End If
       If inout = "0" Then
          If Format(a_time, "HH:MM:SS") < attnrs("intime") Then
             attnrs("intime") = Format(a_time, "HH:MM:SS")
             attnrs("timein") = a_time
             attnrs("indate") = a_date
          End If
       Else
'          If Format(a_time, "HH:MM:SS") > attnrs("intime") Then
             attnrs("outtime") = Format(a_time, "HH:MM:SS")
             attnrs("timeout") = a_time
             attnrs("outdate") = out_date
'          End If
       End If
       attnrs.UPDATE
End Sub
