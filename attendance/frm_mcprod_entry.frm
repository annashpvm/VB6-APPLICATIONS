VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frm_mcprod_entry 
   Caption         =   "M/C PRODUCTION ENTRY"
   ClientHeight    =   9465
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13950
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9465
   ScaleWidth      =   13950
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame3 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   8640
      TabIndex        =   12
      Top             =   3720
      Width           =   2175
      Begin VB.CommandButton save 
         BackColor       =   &H00C0E0FF&
         Caption         =   "&Save"
         Height          =   705
         Left            =   360
         MaskColor       =   &H000000FF&
         Picture         =   "frm_mcprod_entry.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton exit 
         BackColor       =   &H00C0E0FF&
         Caption         =   "&Exit"
         Height          =   705
         Left            =   1080
         MaskColor       =   &H000000FF&
         Picture         =   "frm_mcprod_entry.frx":066A
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   120
         Width           =   735
      End
   End
   Begin VB.Frame Frame4 
      Height          =   975
      Left            =   360
      TabIndex        =   7
      Top             =   9840
      Visible         =   0   'False
      Width           =   3735
      Begin MSComCtl2.DTPicker st_date 
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   61603841
         CurrentDate     =   39359
      End
      Begin MSComCtl2.DTPicker end_date 
         Height          =   375
         Left            =   1920
         TabIndex        =   9
         Top             =   480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   61603841
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
         TabIndex        =   11
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
         TabIndex        =   10
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "M/C PRODUCTION"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8055
      Left            =   1080
      TabIndex        =   5
      Top             =   1440
      Width           =   7455
      Begin MSFlexGridLib.MSFlexGrid flx_data 
         Height          =   6255
         Left            =   480
         TabIndex        =   6
         Top             =   600
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   11033
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Frame Frame5 
      Height          =   615
      Left            =   1560
      TabIndex        =   0
      Top             =   600
      Width           =   6255
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
         TabIndex        =   2
         Top             =   120
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
         TabIndex        =   1
         Top             =   120
         Width           =   3015
      End
      Begin VB.Label Label1 
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
         TabIndex        =   4
         Top             =   240
         Width           =   855
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
         TabIndex        =   3
         Top             =   240
         Width           =   555
      End
   End
End
Attribute VB_Name = "frm_mcprod_entry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

Private Sub exit_Click()
    Unload Me
End Sub

Private Sub flx_data_KeyPress(KeyAscii As Integer)
 On Error GoTo err_handler

 Dim fin_selrow%, fin_selcol%
 fin_selrow = flx_data.Row
 fin_selcol = flx_data.Col
 With flx_data
    If fin_selcol = 2 Then
        If KeyAscii <> 13 Then
            KeyAscii = Numeric_Chk(KeyAscii, flx_data.TextMatrix(fin_selrow, fin_selcol), 8, 5, 3)
            If KeyAscii <> 0 And KeyAscii <> 8 And KeyAscii <> 13 Then
                flx_data.TextMatrix(fin_selrow, fin_selcol) = flx_data.TextMatrix(fin_selrow, fin_selcol) & Chr(KeyAscii)
            ElseIf KeyAscii = 8 Then
              If Len(.TextMatrix(fin_selrow, fin_selcol)) > 0 Then .TextMatrix(fin_selrow, fin_selcol) = Mid(.TextMatrix(fin_selrow, fin_selcol), 1, Len(.TextMatrix(fin_selrow, fin_selcol)) - 1)
              KeyAscii = 0
            End If
        End If
    

    End If
End With
Exit Sub
err_handler:
        chk = gen_Validation(Err.Number, Err.Description)
        If chk = 1 Then
            Resume
        End If

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
    fillgrid
End Sub
Private Sub fillgrid()
   With flx_data
     .Redraw = False
     .Clear
     .Rows = 2
     .Cols = 3
     .TextMatrix(0, 0) = "S.No"
     .TextMatrix(0, 1) = " Date"
     .TextMatrix(0, 2) = "Qty(T)"
     .ColWidth(0) = 500
     .ColWidth(1) = 1600
     .ColWidth(2) = 1200
     
     .Redraw = True
   End With
End Sub

Public Sub find_dates()
     
    If cmb_month.ListIndex = -1 Then Exit Sub
    Dim d1 As Date
    mmon = cmb_month.ItemData(cmb_month.ListIndex)
''    mmon = 10
    
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
    flex_refresh
End Sub

Public Sub flex_refresh()
    fillgrid
    Dim rdate As Date
  
    For rdate = st_date To end_date
         findrow = flx_data.Rows - 1
        flx_data.TextMatrix(findrow, 0) = findrow
        flx_data.TextMatrix(findrow, 1) = rdate
        flx_data.Rows = flx_data.Rows + 1
    Next
    
    
  Dim payrs2 As New ADODB.Recordset

    pst_qry = "select *  from daily_mcprod where mc_date between '" & Format(st_date.Value, "MM/dd/yyyy") & "' and '" & Format(end_date.Value, "MM/dd/yyyy") & "'order by mc_date"
    payrs2.Open pst_qry, paydb, adOpenDynamic, adLockOptimistic
     While Not payrs2.EOF
           For i = 1 To flx_data.Rows - 1
               With flx_data
                    If .TextMatrix(i, 1) = Format(payrs2("mc_date"), "dd/MM/yyyy") Then
                       ''.TextMatrix(i, 2) = payrs2("emps_shift_alloted")
                       .TextMatrix(i, 2) = payrs2("mc_prodn")
                    End If
''                    i = i + 1
               End With
           Next
          payrs2.MoveNext
        Wend
   payrs2.Close

    
End Sub

Private Sub save_Click()
   If flx_data.Rows - 1 = 1 Then Exit Sub
 
paydb.BeginTrans
On Error GoTo err_handler
    Dim pst_qry As String
    Dim payrs As New ADODB.Recordset
    
    pst_qry = "delete from daily_mcprod where mc_date between '" & Format(st_date.Value, "MM/dd/yyyy") & "' and '" & Format(end_date.Value, "MM/dd/yyyy") & "'"
    paydb.Execute pst_qry
    
    pst_qry = "select * from daily_mcprod"
    payrs.Open pst_qry, paydb, 1, 2
    Dim rdate, sdate, edate As Date
    Dim i As Integer
    
    For i = 1 To flx_data.Rows - 1
        sdate = Format(flx_data.TextMatrix(i, 1), "dd/MM/yyyy")
        sql = "insert into daily_mcprod ( mc_date,mc_prodn) values ( '" & Format(sdate, "MM/dd/yyyy") & "','" & Val(flx_data.TextMatrix(i, 2)) & "')"
        paydb.Execute sql
    Next
    
    paydb.CommitTrans
    MsgBox "Record Saved", vbOKOnly + vbInformation, "Information"
    fillgrid
    Exit Sub
Exit Sub
err_handler:
        paydb.RollbackTrans
        chk = gen_Validation(Err.Number, Err.Description)

End Sub
