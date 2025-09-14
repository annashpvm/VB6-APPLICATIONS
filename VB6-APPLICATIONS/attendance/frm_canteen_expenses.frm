VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frm_canteen_expenses 
   Caption         =   "CANTEEN EXPENSES ENTRY"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10935
   ScaleWidth      =   20160
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmd_exit 
      Caption         =   "EXIT"
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
      Left            =   14400
      TabIndex        =   48
      Top             =   6120
      Width           =   1455
   End
   Begin VB.Frame Frame7 
      Height          =   975
      Left            =   14640
      TabIndex        =   6
      Top             =   4920
      Visible         =   0   'False
      Width           =   3735
      Begin MSComCtl2.DTPicker st_date 
         Height          =   375
         Left            =   360
         TabIndex        =   7
         Top             =   480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   119734273
         CurrentDate     =   39359
      End
      Begin MSComCtl2.DTPicker end_date 
         Height          =   375
         Left            =   1920
         TabIndex        =   8
         Top             =   480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   119734273
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
         TabIndex        =   14
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
         TabIndex        =   9
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Frame Frame1 
      Height          =   9495
      Left            =   720
      TabIndex        =   0
      Top             =   240
      Width           =   13575
      Begin MSFlexGridLib.MSFlexGrid flx_data 
         Height          =   3015
         Left            =   240
         TabIndex        =   32
         Top             =   5640
         Width           =   12735
         _ExtentX        =   22463
         _ExtentY        =   5318
         _Version        =   393216
      End
      Begin VB.Frame Frame8 
         Caption         =   "RATE"
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
         Height          =   1215
         Left            =   2040
         TabIndex        =   10
         Top             =   840
         Width           =   9015
         Begin VB.TextBox txt_rate_tea 
            Height          =   375
            Left            =   6600
            TabIndex        =   13
            Top             =   600
            Width           =   1335
         End
         Begin VB.TextBox txt_rate_dinner 
            Height          =   375
            Left            =   4920
            TabIndex        =   12
            Top             =   600
            Width           =   1335
         End
         Begin VB.TextBox txt_rate_lunch 
            Height          =   375
            Left            =   2760
            TabIndex        =   11
            Top             =   600
            Width           =   1335
         End
         Begin VB.TextBox txt_rate_bfast 
            Height          =   375
            Left            =   1080
            TabIndex        =   42
            Top             =   600
            Width           =   1335
         End
         Begin VB.Label Label17 
            Caption         =   "Tea"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   6600
            TabIndex        =   46
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label16 
            Caption         =   "Dinner"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   4920
            TabIndex        =   45
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label15 
            Caption         =   "Breakfast"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   1080
            TabIndex        =   44
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label14 
            Caption         =   "Lunch"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   3240
            TabIndex        =   43
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.Frame Frame2 
         Height          =   3255
         Left            =   120
         TabIndex        =   15
         Top             =   2160
         Width           =   12855
         Begin VB.TextBox txt_others 
            Height          =   375
            Left            =   11280
            TabIndex        =   49
            Top             =   1560
            Width           =   1335
         End
         Begin VB.TextBox txt_snacks 
            Height          =   375
            Left            =   11280
            TabIndex        =   30
            Top             =   960
            Width           =   1335
         End
         Begin VB.CommandButton cmd_save 
            Caption         =   "SAVE"
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
            Left            =   10920
            TabIndex        =   31
            Top             =   2160
            Width           =   1455
         End
         Begin VB.Frame Frame5 
            Caption         =   "TEA TOKEN"
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
            Height          =   2415
            Left            =   6960
            TabIndex        =   38
            Top             =   600
            Width           =   3255
            Begin VB.TextBox txt_tea_early_morn 
               Height          =   375
               Left            =   1680
               TabIndex        =   51
               Top             =   360
               Width           =   1335
            End
            Begin VB.TextBox txt_tea_night 
               Height          =   375
               Left            =   1680
               TabIndex        =   29
               Top             =   1800
               Width           =   1335
            End
            Begin VB.TextBox txt_tea_morn 
               Height          =   375
               Left            =   1680
               TabIndex        =   27
               Top             =   840
               Width           =   1335
            End
            Begin VB.TextBox txt_tea_even 
               Height          =   375
               Left            =   1680
               TabIndex        =   28
               Top             =   1320
               Width           =   1335
            End
            Begin VB.Label Label20 
               Caption         =   "Early Morning"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   495
               Left            =   600
               TabIndex        =   52
               Top             =   360
               Width           =   975
            End
            Begin VB.Label Label13 
               Caption         =   "Night"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   255
               Left            =   600
               TabIndex        =   41
               Top             =   1920
               Width           =   975
            End
            Begin VB.Label Label12 
               Caption         =   "Morning"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   255
               Left            =   600
               TabIndex        =   40
               Top             =   960
               Width           =   975
            End
            Begin VB.Label Label11 
               Caption         =   "Evening"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   255
               Left            =   600
               TabIndex        =   39
               Top             =   1440
               Width           =   975
            End
         End
         Begin VB.Frame Frame4 
            Caption         =   "MD/MANAGERS FOOD"
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
            Height          =   2175
            Left            =   3600
            TabIndex        =   34
            Top             =   840
            Width           =   3255
            Begin VB.TextBox txt_mm_Lunch 
               Height          =   375
               Left            =   1680
               TabIndex        =   24
               Top             =   960
               Width           =   1335
            End
            Begin VB.TextBox txt_mm_breakfast 
               Height          =   375
               Left            =   1680
               TabIndex        =   23
               Top             =   480
               Width           =   1335
            End
            Begin VB.TextBox txt_mm_Dinner 
               Height          =   375
               Left            =   1680
               TabIndex        =   25
               Top             =   1440
               Width           =   1335
            End
            Begin VB.Label Label8 
               Caption         =   "Lunch"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   255
               Left            =   600
               TabIndex        =   37
               Top             =   1080
               Width           =   975
            End
            Begin VB.Label Label7 
               Caption         =   "Breakfast"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   255
               Left            =   600
               TabIndex        =   36
               Top             =   600
               Width           =   975
            End
            Begin VB.Label Label6 
               Caption         =   "Dinner"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   255
               Left            =   600
               TabIndex        =   35
               Top             =   1560
               Width           =   975
            End
         End
         Begin VB.Frame Frame3 
            Caption         =   "EMPLOYEES FOOD"
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
            Height          =   2175
            Left            =   120
            TabIndex        =   18
            Top             =   840
            Width           =   3255
            Begin VB.TextBox txt_Emp_Dinner 
               Height          =   375
               Left            =   1680
               TabIndex        =   22
               Top             =   1440
               Width           =   1335
            End
            Begin VB.TextBox txt_Emp_breakfast 
               Height          =   375
               Left            =   1680
               TabIndex        =   19
               Top             =   480
               Width           =   1335
            End
            Begin VB.TextBox txt_Emp_Lunch 
               Height          =   375
               Left            =   1680
               TabIndex        =   21
               Top             =   960
               Width           =   1335
            End
            Begin VB.Label Label5 
               Caption         =   "Dinner"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   255
               Left            =   600
               TabIndex        =   33
               Top             =   1560
               Width           =   975
            End
            Begin VB.Label Label1 
               Caption         =   "Breakfast"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   255
               Left            =   600
               TabIndex        =   26
               Top             =   600
               Width           =   975
            End
            Begin VB.Label Label3 
               Caption         =   "Lunch"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   255
               Left            =   600
               TabIndex        =   20
               Top             =   1080
               Width           =   975
            End
         End
         Begin MSComCtl2.DTPicker dt_entry 
            Height          =   375
            Left            =   11040
            TabIndex        =   16
            Top             =   240
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   661
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
            Format          =   119734273
            CurrentDate     =   43861
         End
         Begin VB.Label Label19 
            Caption         =   "Others"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   10440
            TabIndex        =   50
            Top             =   1680
            Width           =   975
         End
         Begin VB.Label Label18 
            Caption         =   "Snacks"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   10440
            TabIndex        =   47
            Top             =   1080
            Width           =   975
         End
         Begin VB.Label lbl 
            Caption         =   "Expenses Date"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   9120
            TabIndex        =   17
            Top             =   240
            Width           =   1575
         End
      End
      Begin VB.Frame Frame6 
         Height          =   615
         Left            =   3600
         TabIndex        =   1
         Top             =   120
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
            TabIndex        =   3
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
            TabIndex        =   2
            Top             =   120
            Width           =   3015
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
            TabIndex        =   5
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
            TabIndex        =   4
            Top             =   240
            Width           =   555
         End
      End
   End
End
Attribute VB_Name = "frm_canteen_expenses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmb_month_Click()
    find_dates
End Sub

Private Sub cmb_year_Click()
    find_dates
End Sub

Private Sub cmd_exit_Click()
    Unload Me
End Sub

Private Sub cmd_save_Click()
   
paydb.BeginTrans
On Error GoTo err_handler
    
   
    
     sql = "delete from canteen_expenses where ce_date  = '" & Format(dt_entry, "yyyy/MM/dd") & "'"
     paydb.Execute sql
     sql = "insert into canteen_expenses (ce_date,ce_emp_bfast,ce_emp_lunch,ce_emp_dinner,ce_mm_bfast,ce_mm_lunch,ce_mm_dinner, ce_tea_earlymorn, ce_tea_morn,ce_tea_even,ce_tea_night,ce_rate_bfast,ce_rate_lunch,ce_rate_dinner,ce_rate_tea,ce_snacks,ce_others) " _
         & " values ('" & Format(dt_entry.Value, "MM/dd/yyyy") & "', " & Val(txt_Emp_breakfast.Text) & ", " & Val(txt_Emp_Lunch.Text) & "," & Val(txt_Emp_Dinner.Text) & ", " & Val(txt_mm_breakfast.Text) & ", " & Val(txt_mm_Lunch.Text) & "," & Val(txt_mm_Dinner.Text) & "," & Val(txt_tea_early_morn.Text) & "," & Val(txt_tea_morn.Text) & ", " & Val(txt_tea_even.Text) & "," & Val(txt_tea_night.Text) & " ," & Val(txt_rate_bfast.Text) & "  ," & Val(txt_rate_lunch.Text) & "  ," & Val(txt_rate_dinner.Text) & "  ," & Val(txt_rate_tea.Text) & "," & Val(txt_snacks.Text) & "," & Val(txt_others.Text) & ")"
     paydb.Execute sql
    paydb.CommitTrans
''    MsgBox "Record Saved", vbOKOnly + vbInformation, "Information"
    MsgBox "Record Saved ", vbOKOnly + vbInformation, "Information"""
    refresh1
''    Refresh_Click
    Exit Sub
Exit Sub
err_handler:
        paydb.RollbackTrans
        chk = gen_Validation(Err.Number, Err.Description)

End Sub

Sub refresh1()
    fillgrid
    getdata
    txt_Emp_breakfast.Text = ""
    txt_Emp_Lunch.Text = ""
    txt_Emp_Dinner.Text = ""
    txt_mm_breakfast.Text = ""
    txt_mm_Lunch.Text = ""
    txt_mm_Dinner.Text = ""
    txt_tea_morn.Text = ""
    txt_tea_even.Text = ""
    txt_tea_night.Text = ""
    txt_snacks.Text = ""
    txt_tea_early_morn.Text = ""
    


End Sub

Private Sub flx_data_DblClick()
    Dim pst_ans As String
    If flx_data.Rows - 1 < 1 Then Exit Sub
    fin_selrow = flx_data.Row
    pst_ans = MsgBox("Press YES-to Modify NO-to Cancel", vbYesNoCancel, "Confirmation")
        If pst_ans = vbYes Then
        dt_entry.Value = flx_data.TextMatrix(fin_selrow, 1)
        txt_Emp_breakfast.Text = flx_data.TextMatrix(fin_selrow, 2)
        txt_Emp_Lunch.Text = flx_data.TextMatrix(fin_selrow, 3)
        txt_Emp_Dinner.Text = flx_data.TextMatrix(fin_selrow, 4)
        txt_mm_breakfast.Text = flx_data.TextMatrix(fin_selrow, 6)
        txt_mm_Lunch.Text = flx_data.TextMatrix(fin_selrow, 7)
        txt_mm_Dinner.Text = flx_data.TextMatrix(fin_selrow, 8)
        txt_tea_early_morn.Text = flx_data.TextMatrix(fin_selrow, 10)
        txt_tea_morn.Text = flx_data.TextMatrix(fin_selrow, 11)
        txt_tea_even.Text = flx_data.TextMatrix(fin_selrow, 12)
        txt_tea_night.Text = flx_data.TextMatrix(fin_selrow, 13)
        txt_snacks.Text = flx_data.TextMatrix(fin_selrow, 16)
        
        flx_edit_row = fin_selrow
    ElseIf pst_ans = vbNo Then
        If flx_data.Rows <= 2 Then
           flx_data.Rows = 1
           MsgBox "No Rows to remove"
        Else
           flx_data.RemoveItem fin_selrow
           flx_data.Row = flx_data.Rows - 1
        End If
    End If
End Sub

Private Sub Form_Load()
    dt_entry.Value = Now
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
    
    txt_rate_bfast.Text = "40"
    txt_rate_lunch.Text = "55"
    txt_rate_dinner.Text = "40"
    txt_rate_tea.Text = "6"
    fillgrid
    getdata
End Sub

Private Sub fillgrid()
   With flx_data
     .Redraw = False
     .Clear
     .Rows = 2
     .Cols = 18
     .TextMatrix(0, 0) = "S.No"
     .TextMatrix(0, 1) = "Date"
     .TextMatrix(0, 2) = "Breakfast"
     .TextMatrix(0, 3) = "Lunch"
     .TextMatrix(0, 4) = "Dinner"
     .TextMatrix(0, 5) = "Amount"
     .TextMatrix(0, 6) = "Breakfast"
     .TextMatrix(0, 7) = "Lunch"
     .TextMatrix(0, 8) = "Dinner"
     .TextMatrix(0, 9) = "Amount"
     .TextMatrix(0, 10) = "Ear.Morn"
     .TextMatrix(0, 11) = "Morning"
     .TextMatrix(0, 12) = "Evening"
     .TextMatrix(0, 13) = "Night"
     .TextMatrix(0, 14) = "Total"
     .TextMatrix(0, 15) = "Amount"
     .TextMatrix(0, 16) = "Snacks"
     .TextMatrix(0, 17) = "Others"
     
     .ColWidth(0) = 500
     .ColWidth(1) = 1000
     .ColWidth(2) = 1000
     .ColWidth(3) = 1000
     .ColWidth(4) = 1000
     .ColWidth(5) = 1000
     .ColWidth(6) = 1000
     .ColWidth(7) = 1000
     .ColWidth(8) = 1000
     .ColWidth(9) = 1000
     .ColWidth(10) = 1000
     .ColWidth(11) = 1000
     .ColWidth(12) = 1000
     .ColWidth(13) = 1000
     .ColWidth(14) = 1000
     .ColWidth(15) = 1000
     .ColWidth(16) = 1000
     .ColWidth(17) = 1000
     .Redraw = True
     
   End With

End Sub

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
    
    end_date = DateValue(Str(mmon) + "/" + Str(mdays) + "/" + cmb_year.Text)
    st_date = end_date - Day(end_date) + 1
    getdata

''
''    If end_date.Value > Now Then
''       end_date.Value = Now + 1
''    End If
End Sub

Function getdata()

    fillgrid
    Dim payrs As New ADODB.Recordset
    pst_qry = "select * from canteen_expenses where ce_date between '" & Format(st_date, "MM/dd/yyyy") & "' and '" & Format(end_date, "MM/dd/yyyy") & "'  order by ce_date"
    payrs.Open pst_qry, paydb, adOpenDynamic, adLockOptimistic
    i = 1
    While Not payrs.EOF
        flx_data.TextMatrix(i, 0) = i
        flx_data.TextMatrix(i, 1) = Format(payrs!ce_date, "dd/MM/yyyy")
        flx_data.TextMatrix(i, 2) = payrs!ce_emp_bfast
        flx_data.TextMatrix(i, 3) = payrs!ce_emp_lunch
        flx_data.TextMatrix(i, 4) = payrs!ce_emp_dinner
        flx_data.TextMatrix(i, 5) = (payrs!ce_emp_bfast * payrs!ce_rate_bfast) + (payrs!ce_emp_lunch * payrs!ce_rate_lunch) + (payrs!ce_emp_dinner * payrs!ce_rate_dinner)
        flx_data.TextMatrix(i, 6) = payrs!ce_mm_bfast
        flx_data.TextMatrix(i, 7) = payrs!ce_mm_lunch
        flx_data.TextMatrix(i, 8) = payrs!ce_mm_dinner
        flx_data.TextMatrix(i, 9) = (payrs!ce_mm_bfast * payrs!ce_rate_bfast) + (payrs!ce_mm_lunch * payrs!ce_rate_lunch) + (payrs!ce_mm_dinner * payrs!ce_rate_dinner)
        flx_data.TextMatrix(i, 10) = payrs!ce_tea_earlymorn
        flx_data.TextMatrix(i, 11) = payrs!ce_tea_morn
        flx_data.TextMatrix(i, 12) = payrs!ce_tea_even
        flx_data.TextMatrix(i, 13) = payrs!ce_tea_night
        flx_data.TextMatrix(i, 14) = payrs!ce_tea_earlymorn + payrs!ce_tea_morn + payrs!ce_tea_even + payrs!ce_tea_night
        flx_data.TextMatrix(i, 15) = Val(flx_data.TextMatrix(i, 13)) * payrs!ce_rate_tea
        flx_data.TextMatrix(i, 16) = payrs!ce_snacks
        flx_data.TextMatrix(i, 17) = payrs!ce_others
       
        flx_data.Rows = flx_data.Rows + 1
        

        payrs.MoveNext
        
        i = i + 1
    Wend
    payrs.Close
''  flx_data.Enabled = False

End Function

Private Sub Text1_Change()

End Sub
