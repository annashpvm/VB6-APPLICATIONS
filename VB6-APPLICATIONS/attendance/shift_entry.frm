VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form shift_entry 
   Caption         =   "SHIFT ENTRY"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.ComboBox year_cmb 
      Height          =   315
      Left            =   7440
      TabIndex        =   10
      Top             =   840
      Width           =   1815
   End
   Begin VB.CommandButton exit 
      Caption         =   "&Exit"
      Height          =   825
      Left            =   5385
      Picture         =   "shift_entry.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6855
      Width           =   975
   End
   Begin VB.CommandButton Refresh 
      Caption         =   "&Refresh"
      Height          =   825
      Left            =   4395
      Picture         =   "shift_entry.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6840
      Width           =   975
   End
   Begin VB.CommandButton edit 
      Caption         =   "&Edit"
      Height          =   825
      Left            =   2430
      Picture         =   "shift_entry.frx":0AAC
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6855
      Width           =   975
   End
   Begin VB.CommandButton save 
      Caption         =   "&Save"
      Height          =   825
      Left            =   3360
      Picture         =   "shift_entry.frx":1116
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6855
      Width           =   975
   End
   Begin VB.ComboBox employee_cmb 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3480
      TabIndex        =   4
      Top             =   1680
      Width           =   5775
   End
   Begin VB.ComboBox month_cmb 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3480
      TabIndex        =   2
      Top             =   840
      Width           =   2895
   End
   Begin MSFlexGridLib.MSFlexGrid sft_flex 
      Height          =   4455
      Left            =   1560
      TabIndex        =   0
      Top             =   2280
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   7858
      _Version        =   393216
   End
   Begin VB.Label Label3 
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
      Height          =   495
      Left            =   6600
      TabIndex        =   9
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Select Employee"
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
      Left            =   1560
      TabIndex        =   3
      Top             =   1800
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Month"
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
      Left            =   1560
      TabIndex        =   1
      Top             =   960
      Width           =   1215
   End
End
Attribute VB_Name = "shift_entry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim attndb As New ADODB.Connection
Dim attnrs As New ADODB.Recordset
Dim endrow As Integer
Private Sub employee_cmb_Click()
        fillgrid
        sdate = Date
        edate = Date
        If month_cmb.Text = "APRIL" Then
           sdate = CDate("01/04/2007")
           edate = CDate("30/04/2007")
        ElseIf month_cmb.Text = "MAY" Then
           sdate = CDate("01/05/2007")
           edate = CDate("31/05/2007")
        ElseIf month_cmb.Text = "JUNE" Then
           sdate = CDate("01/06/2007")
           edate = CDate("30/06/2007")
        ElseIf month_cmb.Text = "JULY" Then
           sdate = CDate("01/07/2007")
           edate = CDate("31/07/2007")
        ElseIf month_cmb.Text = "AUGUEST" Then
           sdate = CDate("01/08/2007")
           edate = CDate("31/08/2007")
        ElseIf month_cmb.Text = "SEPTEMBER" Then
           sdate = CDate("01/09/2007")
           edate = CDate("30/09/2007")
        ElseIf month_cmb.Text = "OCTOBER" Then
           sdate = CDate("01/10/2007")
           edate = CDate("31/10/2007")
        ElseIf month_cmb.Text = "NOVEMBER" Then
           sdate = CDate("01/11/2007")
           edate = CDate("30/11/2007")
        ElseIf month_cmb.Text = "DECEMBER" Then
           sdate = CDate("01/12/2007")
           edate = CDate("31/12/2007")
        ElseIf month_cmb.Text = "JANUARY" Then
           sdate = CDate("01/01/2008")
           edate = CDate("31/01/2008")
        ElseIf month_cmb.Text = "FEBRUARY" Then
           sdate = CDate("01/02/2008")
           edate = CDate("28/02/2008")
        ElseIf month_cmb.Text = "MARCH" Then
           sdate = CDate("01/03/2008")
           edate = CDate("31/03/2008")
        Else
            MsgBox ("Select a month")
            Exit Sub
        End If
              rdate = sdate
           While Not rdate > edate
               With sft_flex
                 .Rows = .Rows + 1
                 .TextMatrix(.Rows - 1, 0) = rdate
                 .TextMatrix(.Rows - 1, 1) = "N"
                 rdate = rdate + 1
                 endrow = endrow + 1
               End With
            Wend
        attndb.Open "Provider=SQLOLEDB.1;Password=serdat;Persist Security Info=True;User ID=sa;DATABASE=servall;Data Source=servalldata"
        Dim attnrs As New ADODB.Recordset
        sql_qry = "select * from attn_emp_shift where semp_name = '" & Trim(employee_cmb.Text) & "' and semp_month = '" & month_cmb.Text & "' and semp_year = '" & year_cmb.Text & "' order by semp_date"
        attnrs.Open sql_qry, attndb, adOpenDynamic, adLockOptimistic
        While Not attnrs.EOF
               For i = 1 To endrow - 1
                  If sft_flex.TextMatrix(i, 0) = attnrs("semp_date") Then
                     sft_flex.TextMatrix(i, 1) = "Y"
                  End If
               Next
               attnrs.MoveNext
        Wend
        attndb.Close
        
        
End Sub

Private Sub EXIT_Click()
        Unload Me
End Sub

Private Sub Form_Load()
        dsnmdb = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=\\personnel\plat04s\Acu301.mdb"
'        dsnmdb = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\to be deleted\Acu301.mdb"
        month_cmb.AddItem ("APRIL")
        month_cmb.AddItem ("MAY")
        month_cmb.AddItem ("JUNE")
        month_cmb.AddItem ("JULY")
        month_cmb.AddItem ("AUGUEST")
        month_cmb.AddItem ("SEPTEMBER")
        month_cmb.AddItem ("OCTOBER")
        month_cmb.AddItem ("NOVEMBER")
        month_cmb.AddItem ("DECEMBER")
        month_cmb.AddItem ("JANUARY")
        month_cmb.AddItem ("FEBRUARY")
        month_cmb.AddItem ("MARCH")
        fillgrid
        Dim mdbrs As New ADODB.Recordset
        mdb_qry = "Select * from Member where companyID <> '9999'"
        mdbrs.Open mdb_qry, dsnmdb, adOpenDynamic, adLockOptimistic
        totemp = 0
        While Not mdbrs.EOF
             employee_cmb.AddItem (mdbrs(1))
             employee_cmb.ItemData(employee_cmb.NewIndex) = mdbrs(0)
             mdbrs.MoveNext
        Wend
        For i = 2007 To 2020
            year_cmb.AddItem (i)
        Next
End Sub

Function fillgrid()
    With sft_flex
        .Clear
        .Cols = 2
        .Rows = 1
        .TextMatrix(0, 0) = "DATE"
        .TextMatrix(0, 1) = "C shift (Y/N)"
        .ColWidth(0) = 2000
        .ColWidth(1) = 3500
    End With
End Function

Private Sub Refresh_Click()
     fillgrid
End Sub

Private Sub SAVE_Click()
    If month_cmb.Text = "" Then
     MsgBox ("Select the Month")
     Exit Sub
   End If
   If year_cmb.Text = "" Then
     MsgBox ("Select the Year")
     Exit Sub
   End If
   If employee_cmb.Text = "" Then
     MsgBox ("Select any Employee")
     Exit Sub
   End If
   attndb.Open "Provider=SQLOLEDB.1;Password=serdat;Persist Security Info=True;User ID=sa;DATABASE=servall;Data Source=servalldata"
   Set attnrs = New ADODB.Recordset
   sql_qry = "delete from attn_emp_shift where semp_name = '" & Trim(employee_cmb.Text) & "' and semp_month = '" & month_cmb.Text & "' and semp_year = '" & year_cmb.Text & "'"
   attnrs.Open sql_qry, attndb, adOpenDynamic, adLockOptimistic
   
   Set attnrs = New ADODB.Recordset
   sql_qry = "select * from attn_emp_shift"
   attnrs.Open sql_qry, attndb, adOpenDynamic, adLockOptimistic
   
   rec_chk = 0
   For i = 1 To endrow
       If Trim(sft_flex.TextMatrix(i, 1)) = "Y" Then
            attnrs.AddNew
            attnrs.Fields("semp_code") = employee_cmb.ItemData(employee_cmb.ListIndex)
            attnrs.Fields("semp_name") = employee_cmb.Text
            attnrs.Fields("semp_month") = month_cmb.Text
            attnrs.Fields("semp_year") = Val(year_cmb.Text)
            attnrs.Fields("semp_date") = CDate(sft_flex.TextMatrix(i, 0))
            attnrs.Fields("semp_cshift") = Trim(sft_flex.TextMatrix(i, 1))
            rec_chk = 1
            attnrs.Update
       End If
   Next
   endrow = 0
   If rec_chk = 1 Then
     MsgBox ("Records are saved")
     fillgrid
  Else
     MsgBox ("Details not saved")
  End If
  attndb.Close
End Sub

Private Sub sft_flex_KeyPress(KeyAscii As Integer)
On Error GoTo err_handler
 Dim fin_selrow%, fin_selcol%
 fin_selrow = sft_flex.Row
 fin_selcol = sft_flex.Col
 With sft_flex
 Select Case fin_selcol
        Case 1
           If sft_flex.TextMatrix(fin_selrow, fin_selcol) = "N" Then
               sft_flex.TextMatrix(fin_selrow, fin_selcol) = "Y"
           Else
               sft_flex.TextMatrix(fin_selrow, fin_selcol) = "N"
           End If
  
'         If KeyAscii <> 0 And KeyAscii <> 8 And KeyAscii <> 13 Then
 '            sft_flex.TextMatrix(fin_selrow, fin_selcol) = sft_flex.TextMatrix(fin_selrow, fin_selcol) & Chr(KeyAscii)
    End Select
 End With
Exit Sub
err_handler:
''        chk = gen_Validation(Err.Number, Err.Description)
''        If chk = 1 Then
''            Resume
''        End If
End Sub
