VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frm_gratuity_settlement 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport cry_rep 
      Left            =   1560
      Top             =   8760
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame Frame3 
      Height          =   975
      Left            =   4680
      TabIndex        =   5
      Top             =   6600
      Width           =   1695
      Begin VB.CommandButton PROCESS 
         Caption         =   "&PRINT"
         Height          =   825
         Left            =   120
         Picture         =   "frm_gratuity_settlement.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   120
         Width           =   720
      End
      Begin VB.CommandButton EXIT 
         Caption         =   "E&XIT"
         Height          =   825
         Left            =   840
         Picture         =   "frm_gratuity_settlement.frx":066A
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   120
         Width           =   720
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " "
      Height          =   7335
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   11175
      Begin VB.CommandButton cmd_process 
         Caption         =   "Process"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   10200
         TabIndex        =   13
         Top             =   1560
         Width           =   975
      End
      Begin VB.ComboBox cmb_dept 
         Height          =   315
         Left            =   3360
         TabIndex        =   11
         Text            =   " "
         Top             =   720
         Width           =   4335
      End
      Begin VB.Frame Frame2 
         Height          =   615
         Left            =   3480
         TabIndex        =   8
         Top             =   0
         Width           =   3255
         Begin VB.OptionButton Opt_worker 
            Caption         =   "Worker"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   375
            Left            =   1680
            TabIndex        =   10
            Top             =   120
            Width           =   1335
         End
         Begin VB.OptionButton Opt_Staff 
            Caption         =   "Staff"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   375
            Left            =   120
            TabIndex        =   9
            Top             =   120
            Width           =   1935
         End
      End
      Begin MSFlexGridLib.MSFlexGrid flx_data 
         Height          =   4095
         Left            =   480
         TabIndex        =   4
         Top             =   2040
         Width           =   9615
         _ExtentX        =   16960
         _ExtentY        =   7223
         _Version        =   393216
      End
      Begin VB.ComboBox cmb_name 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3360
         TabIndex        =   3
         Text            =   " "
         Top             =   1200
         Width           =   6855
      End
      Begin VB.Label Label4 
         Caption         =   "Department"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   960
         TabIndex        =   12
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label Label2 
         Caption         =   "Employee Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   960
         TabIndex        =   2
         Top             =   1200
         Width           =   1815
      End
   End
   Begin VB.Label Label1 
      Caption         =   "GRATUITY SETTLEMENT"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   375
      Left            =   2640
      TabIndex        =   1
      Top             =   120
      Width           =   5175
   End
End
Attribute VB_Name = "frm_gratuity_settlement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmb_dept_Click()
        cmb_name.Clear
        Set paydb = New ADODB.Connection
        Set payrs = New ADODB.Recordset
        
        If opt_staff.Value = True Then
            sql = "select * from emp_mas  where emp_company = '" & company_code & "'  and emp_cat in ('S','M') and emp_dept =' " & cmb_dept.ItemData(cmb_dept.ListIndex) & " ' order by emp_name"
        ElseIf opt_worker.Value = True Then
             sql = "select * from emp_mas  where emp_company = '" & company_code & "'  and emp_cat in ('W') and emp_dept =' " & cmb_dept.ItemData(cmb_dept.ListIndex) & " ' order by emp_name"
        End If
        
        paydb.Open pay
        payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
        If Not payrs.EOF Then
            payrs.MoveFirst
            cmb_name.Clear
            While Not payrs.EOF
                cmb_name.AddItem payrs("emp_name")
                payrs.MoveNext
            Wend
        End If
End Sub

Private Sub cmd_process_Click()
    
    If Trim(cmb_dept.Text) = "" Then
        MsgBox " Select Department to proceed further "
        cmb_dept.SetFocus
        Exit Sub
    End If
    If Trim(cmb_name.Text) = "" Then
        MsgBox " Select Name of the Employee  "
        cmb_name.SetFocus
        Exit Sub
    End If
    
    Set paydb = New ADODB.Connection
    Set payrs = New ADODB.Recordset
    
    sql = ("Select pdesi_name,emp_doj,emp_resigneddate,emp_basic,emp_splpay,emp_cat,emp_serwt,emp_fda,emp_vda from emp_mas a,pdesi_mas b where emp_name='" & cmb_name.Text & "' and a.emp_design=b.pdesi_code")
''    sql = "Select pdesi_name,emp_doj,emp_resigneddate,s_basic as emp_basic,s_splpay as emp_splpay,emp_cat,s_serwt as emp_serwt,s_fda as emp_fda,s_vda as emp_vda  from emp_mas a,pdesi_mas b,emp_salary c where emp_code=s_empcode and emp_name='" & cmb_name.Text & "' and a.emp_design=b.pdesi_code and s_month =month(a.emp_resigneddate) and s_company=" & company_code & " and s_finyear=" & finyear & ""
    paydb.Open pay
    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
    If payrs.EOF = False Then
    If IsNull(payrs!emp_resigneddate) Then
        MsgBox "EMPLOYEE NOT RESIGNED/RESIGNED DATE NOT ENTERED IN EMPLOYEE MASTER"
        Exit Sub
    End If
    Else
        MsgBox "MISSING VALUES"
        Exit Sub
    End If
    payrs.Close
    sql = "Select pdesi_name,emp_doj,emp_resigneddate,s_basic as emp_basic,s_splpay as emp_splpay,emp_cat,s_serwt as emp_serwt,s_fda as emp_fda,s_vda as emp_vda  from emp_mas a,pdesi_mas b,emp_salary c where emp_code=s_empcode and emp_name='" & cmb_name.Text & "' and a.emp_design=b.pdesi_code and s_month =month(a.emp_resigneddate) and s_company=" & company_code & " and s_finyear=" & finyear & ""
    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
    If payrs.EOF = True Then
        MsgBox "GO TO PREVIOUS FINANCIAL YEARS TO GET HIS/HER GRATUITY "
        Exit Sub
    End If
    Dim Month, year, Basic As Integer
    Dim monthname, yearname As String
    monthname = "  Months "
    yearname = "  Years "
    Month = DateDiff("M", payrs!emp_doj, payrs!emp_resigneddate)
    
    If payrs!emp_cat = "W" Then
        Basic = payrs!emp_basic + payrs!emp_serwt + payrs!emp_fda + payrs!emp_fda
    Else
        ''Basic = payrs!emp_basic + payrs!emp_splpay
        Basic = payrs!emp_basic
    End If
    
    Do Until Month < 12
    
        year = year + 1
        Month = Month - 12
    
    Loop
    With flx_data
        flx_data.TextMatrix(1, 1) = payrs!pdesi_name
        flx_data.TextMatrix(2, 1) = Format(payrs!emp_doj, "dd/MM/yyyy")
        flx_data.TextMatrix(3, 1) = Format(payrs!emp_resigneddate, "dd/MM/yyyy")
        If year > 0 And Month > 0 Then
            flx_data.TextMatrix(4, 1) = year & yearname & Month & monthname
        ElseIf year = 0 And Month > 0 Then
            flx_data.TextMatrix(4, 1) = Month & monthname
        ElseIf year > 0 And Month = 0 Then
            flx_data.TextMatrix(4, 1) = year & yearname
        Else
            flx_data.TextMatrix(4, 1) = "NOT ELIGIBLE FOR GRATUITY"
        End If
        flx_data.TextMatrix(5, 1) = Basic
        If payrs!emp_cat = "W" Then
            flx_data.TextMatrix(6, 1) = Round(Basic * (15 / 26))
            flx_data.TextMatrix(7, 1) = Round(year * Basic * 15 / 26)
        Else
            flx_data.TextMatrix(6, 1) = Round(Basic * 4.8 / 100)
            flx_data.TextMatrix(7, 1) = Round(((year * 12) + Month) * (Basic * 4.8) / 100)
        End If
        
        flx_data.ColAlignment(1) = flexAlignLeftCenter
    End With
End Sub

Private Sub exit_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Set paydb = New ADODB.Connection
    Set payrs = New ADODB.Recordset
    sql = ("Select * from  pdept_mas order by dept_name")
    paydb.Open pay
    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
    payrs.MoveFirst
    While Not payrs.EOF
        cmb_dept.AddItem payrs(1)
        cmb_dept.ItemData(cmb_dept.NewIndex) = payrs(0)
        payrs.MoveNext
    Wend
    fillgrid
    opt_staff.Value = True
End Sub

Function fillgrid()
Dim i As Integer
    With flx_data
        .Clear
        .Cols = 2
        .Rows = 8
        .TextMatrix(0, 0) = "DESCRIPTION"
        .TextMatrix(1, 0) = "DESIGNATION"
        .TextMatrix(2, 0) = "DATE OF JOINING"
        .TextMatrix(3, 0) = "DATE OF RESIGNATION "
        .TextMatrix(4, 0) = "NO OF YEARS SERVED"
        .TextMatrix(5, 0) = "BASIC PAY"
        .TextMatrix(6, 0) = "GRATUITY ELIGIBLE PER MONTH"
        .TextMatrix(7, 0) = "GRATUITY AMOUNT"
        .TextMatrix(0, 1) = "VALUES"
        
        For i = 0 To 7
            flx_data.RowHeight(i) = 500
        Next
        
        flx_data.Font.Bold = True
        flx_data.ForeColorFixed = vbBlue
        .ColWidth(0) = 3000
        .ColWidth(1) = 8000
      
    End With
End Function

Private Sub PROCESS_Click()
''On Error GoTo err_handler

cry_rep.ReportFileName = "\\10.0.0.252\vbcryrep\payroll\gratuity_settlement_individual.rpt"
cry_rep.Formulas(1) = ("millname= '" & millname & "'")
cry_rep.Formulas(2) = ("name= '" & cmb_name & "'")
cry_rep.Formulas(3) = ("dept= '" & cmb_dept & "'")
cry_rep.Formulas(4) = ("desi= '" & flx_data.TextMatrix(1, 1) & "'")
cry_rep.Formulas(5) = ("doj= '" & flx_data.TextMatrix(2, 1) & "'")
cry_rep.Formulas(6) = ("dor= '" & flx_data.TextMatrix(3, 1) & "'")
cry_rep.Formulas(7) = ("noy= '" & flx_data.TextMatrix(4, 1) & "'")
cry_rep.Formulas(8) = ("basic= '" & flx_data.TextMatrix(5, 1) & "'")
cry_rep.Formulas(9) = ("eligible= '" & flx_data.TextMatrix(6, 1) & "'")
cry_rep.Formulas(10) = ("amount= '" & flx_data.TextMatrix(7, 1) & "'")


''cry_rep.ReplaceSelectionFormula pst_qry
cry_rep.PrinterSelect
cry_rep.Connect = gst_reportconnect
cry_rep.PageZoom (110)
''cry_rep.WindowMaxButton (110)
cry_rep.Action = 1
End Sub
