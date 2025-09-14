VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frm_upload 
   Caption         =   "DATA UPLOADING"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   5655
      Left            =   1200
      TabIndex        =   0
      Top             =   960
      Width           =   9615
      Begin VB.CommandButton cmd_upload 
         Caption         =   "Upload"
         Height          =   495
         Left            =   4320
         TabIndex        =   4
         Top             =   2520
         Width           =   1875
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "Exit"
         Height          =   495
         Left            =   6360
         TabIndex        =   3
         Top             =   2520
         Width           =   1875
      End
      Begin VB.CommandButton cmd_select 
         Caption         =   "Select file"
         Height          =   495
         Left            =   2160
         TabIndex        =   2
         Top             =   2520
         Width           =   1875
      End
      Begin VB.TextBox txt_filename 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2160
         TabIndex        =   1
         Top             =   720
         Width           =   6135
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   720
         Top             =   1920
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label Label1 
         Caption         =   "FILE PATH"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   840
         Width           =   1695
      End
   End
End
Attribute VB_Name = "frm_upload"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim xl As New Excel.Application
Dim xlsheet As Excel.Worksheet
Dim xlwbook As Excel.Workbook
Dim pst_qry As String
Private Sub cmd_select_Click()
Dim ff As Integer
ff = FreeFile 'Sets to next available file number
With CommonDialog1
    .FileName = ""
''    .Filter = "All files (*.xls) |*.xls|" 'Sets the filter

    .Filter = "All files (*.*)|*.*" 'Sets the filter
''    .Filter = "All files (*.xlsx) |*.xlsx|" 'Sets the filter
    .ShowOpen
End With
txt_filename.Text = CommonDialog1.FileName
End Sub
Private Sub cmd_upload_Click()
Dim col1 As String
''Dim col2 as
Dim col3 As Long

''    Dim payrs As New ADODB.Recordset
    Set xlwbook = xl.Workbooks.Open("" & txt_filename.Text & "")
    Set xlsheet = xlwbook.Sheets.item(1)
    Dim i As Integer
    For i = 1 To 3
  ''      MsgBox (xlsheet.Cells(i, 1) + " - " + Str(xlsheet.Cells(i, 2)))
''        pst_qry = "select * from emp_mas where emp_company = " & company_code & " and emp_cat = 'S' and emp_no = '" & xlsheet.Cells(i, 1) & "'"
''        payrs.Open pst_qry, paydb, adOpenDynamic, adLockOptimistic

      col1 = xlsheet.Cells(i, 1)
      col2 = xlsheet.Cells(i, 2)
      col3 = xlsheet.Cells(i, 3)
      pst_qry = "insert into EXCELTEST values ( '" & col1 & "' , '" & col2 & "' , '" & col3 & "')"
      paydb.Execute pst_qry


    Next
    xl.ActiveWorkbook.Close False, "" & txt_filename.Text & ""
    xl.Quit
    Set xlwbook = Nothing
    Set xl = Nothing
    MsgBox ("Data Upload completed...")
    
''    Dim payrs As New ADODB.Recordset
''    sql = "Select * from  emp_mas where emp_company = '" & company_code & "' and emp_status in  ('B','C')  order by emp_doj"
''    emp_cat = "W"
''    paydb.Open pay
''    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set xlwbook = Nothing
    Set xl = Nothing
End Sub
