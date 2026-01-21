VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frm_txt_to_sql 
   AutoRedraw      =   -1  'True
   Caption         =   "Test"
   ClientHeight    =   5370
   ClientLeft      =   615
   ClientTop       =   1650
   ClientWidth     =   8205
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   5370
   ScaleWidth      =   8205
   Begin VB.TextBox Text6 
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
      Left            =   1440
      TabIndex        =   3
      Top             =   2040
      Width           =   6135
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   2400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Select file"
      Height          =   495
      Left            =   1440
      TabIndex        =   2
      Top             =   3000
      Width           =   1875
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   495
      Left            =   5640
      TabIndex        =   1
      Top             =   3000
      Width           =   1875
   End
   Begin VB.CommandButton cmdTryIt 
      Caption         =   "Upload"
      Height          =   495
      Left            =   3600
      TabIndex        =   0
      Top             =   3000
      Width           =   1875
   End
End
Attribute VB_Name = "frm_txt_to_sql"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
     (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As _
     String, ByVal lpParameters As String, ByVal lpDirectory _
     As String, ByVal nShowCmd As Long) As Long


Private Sub cmdClear_Click()
    Cls
End Sub


Private Sub cmdExit_Click()
    Unload Me
End Sub


Private Sub cmdTryIt_Click()
    Dim pst_qry As String
    Dim em_idcode, em_fpcode, em_code, em_name, em_fname, EM_PFNO, EM_ACNO, EM_DA, EM_PF, EM_MILLCBE, EM_COMPANY, EM_CAT  As String
    Dim EM_BASIC, EM_SWC, EM_HRA, EM_CALW, EM_LIC, EM_RD, EM_MA, EM_SPL, EM_AA, EM_HEALTH, EM_WASH, EM_SPLPAY, EM_MEALS, EM_LTA, EM_EDU, EM_MAG, EM_OTHALW, EM_TEAALW, EM_TEADED, EM_WFUND, EM_PFDED, EM_BKDED As Single
    Dim EM_DOB, EM_DOJ As String
    Dim strEmpFileName  As String
    Dim strBackSlash  As String
    Dim intEmpFileNbr As Integer
    If Text6.Text = "" Then
        MsgBox ("Select Text file...")
        Exit Sub
'       strBackSlash = IIf(Right$(App.Path, 1) = "\", "", "\")
 ''      strEmpFileName = App.Path & strBackSlash & "EMPLOYEE.DAT"
    Else
       strEmpFileName = Text6.Text
    End If
  ''  strEmpFileName = App.Path & strBackSlash & "t.txt"
    
    intEmpFileNbr = FreeFile
    Open strEmpFileName For Input As #intEmpFileNbr
    Dim r As Integer
    r = 1
    Do Until EOF(intEmpFileNbr)
        Input #intEmpFileNbr, em_idcode, em_fpcode, em_code, em_name, em_fname, EM_PFNO, EM_BASIC, EM_SWC, EM_HRA, EM_CALW, EM_LIC, EM_RD, EM_MA, EM_SPL, EM_AA, EM_HEALTH, EM_WASH, EM_SPLPAY, EM_MEALS, EM_LTA, EM_EDU, EM_MAG, EM_OTHALW, EM_TEAALW, EM_TEADED, EM_WFUND, EM_PFDED, EM_BKDED, EM_ACNO, EM_DA, EM_PF, EM_MILLCBE, EM_COMPANY, EM_CAT, EM_DOB, EM_DOJ
''   @ W_LINE,0   SAY LEFT(em_idcode,8)+","+str(em_fpcode,4)+","+left(em_code,6)+","+EM_NAME+","+EM_FNAME+","+STR(EM_PFNO,8)+","+STR(EM_BASIC,8,2)+","+STR(EM_SWC,8,2)+","+STR(EM_HRA,8,2)+","+STR(EM_CALW,8,2)+","+STR(EM_LIC,8,2)+","+STR(EM_RD,8,2) + ","
  '' @ W_LINE,148  SAY STR(EM_MA,8,2)+","+STR(EM_SPL,8,2)+","+STR(EM_AA,8,2) + "," +STR(EM_HEALTH,8,2)+","+STR(EM_WASH,8,2)+","+STR(EM_SPLPAY,8,2)+","+STR(EM_MEALS,8,2)+","+STR(EM_LTA,8,2) +","+STR(EM_EDU,8,2) + ","+STR(EM_MAG,8,2) + ","
   ''@ W_LINE,240  SAY  +","+STR(EM_OTHALW,8,2) + "," +STR(EM_TEAALW,8,2)+","+STR(EM_TEADED,8,2)+","+STR(EM_WFUND,8,2) +","+STR(EM_PFDED,8,2) +","+STR(EM_BKDED,8,2) + ","+ EM_ACNO
        r = r + 1
        If r >= 34 Then
           r = r
        End If
        pst_qry = "insert into emp_mas (EMP_IDCODE , EMP_FPCODE, EMP_CODE, EMP_NAME, EMP_FNAME, EMP_PFNO, EMP_BASIC, EMP_SERWT, EMP_HRA, EMP_CONVALL, EMP_LIC, EMP_RD, EMP_MEDALL, EMP_SPLALL, EMP_ATTALL, EMP_HEALTH, EMP_WASHALL, EMP_SPLPAY, EMP_MEALSALL, EMP_LTA, EMP_EDUALL, EMP_MAGALL, EMP_OTHALL, EMP_TEAALL, EMP_TEADED, EMP_WFUND, EMP_PFDED, EMP_BANKDED, EMP_BANK_ACNO , EMP_DA_ELIGIBLE, EMP_PFELIGIBLE, EMP_WORKPLACE, EMP_COMPANY, EMP_CAT, EMP_DOB, EMP_DOJ)  values ( " & em_idcode & " , " & em_fpcode & ", " & em_code & " , '" & UCase(em_name) & "' , '" & UCase(em_fname) & "','" & EM_PFNO & "'," & EM_BASIC & ", " & EM_SWC & ", " & EM_HRA & ", " & EM_CALW & ", " & EM_LIC & ", " & EM_RD & ", " & EM_MA & ", " & EM_SPL & ", " & EM_AA & "," & EM_HEALTH & ", " & EM_WASH & ", " & EM_SPLPAY & ", " & EM_MEALS & ", " & EM_LTA & ", " & EM_EDU & ", " & EM_MAG & ", " & EM_OTHALW & ", " & EM_TEAALW & " , " & EM_TEADED & ", " & EM_WFUND & ", " & EM_PFDED & ", " & EM_BKDED & ", '" & EM_ACNO & "'," _
                  & "'" & EM_DA & " ' , '" & EM_PF & "' , '" & EM_MILLCBE & "' ,  " & EM_COMPANY & ",   '" & EM_CAT & "' ,  " & Format(EM_DOB, "MM/dd/yyyy") & " ,  " & EM_DOJ & "  ) "
        paydb.Execute pst_qry
    Loop
    Close #intEmpFileNbr
    MsgBox ("Data uploaded...")
End Sub

Private Sub Command1_Click()
Dim ff As Integer
ff = FreeFile 'Sets to next available file number
With CommonDialog1
    .FileName = ""
''    .Filter = "All files (*.*) |*.*|" 'Sets the filter
    .Filter = "All files (*.txt) |*.txt|" 'Sets the filter
    
    .ShowOpen
End With
Text6 = CommonDialog1.FileName
'If CommonDialog1.FileName = "" Then Exit Sub
'Open CommonDialog1.FileName For Input As #ff 'Opens for reading
'    Text6 = Input(LOF(ff), ff) 'Retrieves all data
'Close #ff 'closes file


''ShellExecute 0&, "open", CommonDialog1.FileName, "", "", vbNormalFocus

End Sub

