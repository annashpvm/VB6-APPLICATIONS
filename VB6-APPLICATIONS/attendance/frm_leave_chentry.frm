VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frm_leave_chentry 
   Caption         =   "COMPENSATION ENTRIES"
   ClientHeight    =   8925
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8925
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame3 
      Caption         =   "CH AVAILED FOR"
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
      Height          =   735
      Left            =   6960
      TabIndex        =   43
      Top             =   0
      Width           =   6015
      Begin VB.OptionButton Option3 
         Caption         =   "Declare Holiday Worked"
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
         Left            =   2040
         TabIndex        =   47
         Top             =   240
         Width           =   1695
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Week off Worked"
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
         Left            =   480
         TabIndex        =   45
         Top             =   240
         Value           =   -1  'True
         Width           =   1575
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Extra Hours worked"
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
         Left            =   4200
         TabIndex        =   44
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame Frame7 
      Height          =   975
      Left            =   240
      TabIndex        =   37
      Top             =   7320
      Width           =   3735
      Begin MSComCtl2.DTPicker st_date 
         Height          =   375
         Left            =   360
         TabIndex        =   38
         Top             =   480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   61210625
         CurrentDate     =   39359
      End
      Begin MSComCtl2.DTPicker end_date 
         Height          =   375
         Left            =   1920
         TabIndex        =   39
         Top             =   480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   61210625
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
         TabIndex        =   41
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
         TabIndex        =   40
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame6 
      Height          =   615
      Left            =   1800
      TabIndex        =   32
      Top             =   720
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
         TabIndex        =   34
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
         TabIndex        =   33
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
         Left            =   4920
         TabIndex        =   36
         Top             =   120
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
         TabIndex        =   35
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Height          =   5415
      Left            =   240
      TabIndex        =   6
      Top             =   1320
      Width           =   3015
      Begin VB.CommandButton cmd_clear 
         Caption         =   "CLEAR"
         Height          =   255
         Left            =   1440
         TabIndex        =   12
         Top             =   4920
         Width           =   975
      End
      Begin VB.CommandButton cmd_filter 
         Caption         =   "FILTER"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   4920
         Width           =   975
      End
      Begin VB.ListBox lst_employee 
         Height          =   1425
         Left            =   120
         TabIndex        =   10
         Top             =   3360
         Width           =   2655
      End
      Begin VB.ListBox lst_dept 
         Height          =   1425
         Left            =   120
         TabIndex        =   9
         Top             =   1560
         Width           =   2655
      End
      Begin VB.TextBox txt_empcode 
         Height          =   285
         Left            =   1440
         TabIndex        =   8
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox txt_empname 
         Height          =   285
         Left            =   120
         TabIndex        =   7
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
         TabIndex        =   16
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
         TabIndex        =   15
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
         TabIndex        =   14
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
         TabIndex        =   13
         Top             =   720
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      Height          =   6015
      Left            =   3360
      TabIndex        =   2
      Top             =   1320
      Width           =   9015
      Begin VB.TextBox txt_cat 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   7080
         TabIndex        =   48
         Top             =   360
         Width           =   1575
      End
      Begin VB.CommandButton cmd_Assign 
         Caption         =   "Assign CH"
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
         Left            =   7080
         TabIndex        =   28
         Top             =   3120
         Width           =   1695
      End
      Begin VB.CommandButton cmd_modify 
         Caption         =   "Modify CH"
         Enabled         =   0   'False
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
         Left            =   7080
         TabIndex        =   30
         Top             =   5040
         Width           =   1695
      End
      Begin VB.CommandButton cmd_delete 
         Caption         =   "Delete CH"
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
         Left            =   7080
         TabIndex        =   29
         Top             =   4080
         Width           =   1695
      End
      Begin MSComCtl2.DTPicker dt_date 
         Height          =   255
         Left            =   6960
         TabIndex        =   42
         Top             =   2160
         Visible         =   0   'False
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   450
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   61210625
         CurrentDate     =   42278
      End
      Begin VB.TextBox txt_dept 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   5040
         TabIndex        =   23
         Top             =   360
         Width           =   1815
      End
      Begin VB.TextBox txt_fpcode 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   240
         TabIndex        =   22
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox txt_empname2 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1680
         TabIndex        =   21
         Top             =   360
         Width           =   3135
      End
      Begin VB.Frame frame_leavetype 
         Caption         =   "Leave"
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
         Height          =   735
         Left            =   3480
         TabIndex        =   18
         Top             =   960
         Visible         =   0   'False
         Width           =   3255
         Begin VB.OptionButton opt_an 
            Caption         =   "After Noon"
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
            Left            =   1680
            TabIndex        =   20
            Top             =   360
            Width           =   1335
         End
         Begin VB.OptionButton opt_fn 
            Caption         =   "Fore Noon"
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
            Left            =   120
            TabIndex        =   19
            Top             =   360
            Width           =   1215
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Leave Type"
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
         Height          =   735
         Left            =   240
         TabIndex        =   3
         Top             =   960
         Width           =   3135
         Begin VB.OptionButton opt_leave_full 
            Caption         =   "Full Day"
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
            TabIndex        =   5
            Top             =   360
            Value           =   -1  'True
            Width           =   1215
         End
         Begin VB.OptionButton opt_leave_half 
            Caption         =   "Half Day"
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
            Left            =   1800
            TabIndex        =   4
            Top             =   360
            Width           =   1215
         End
      End
      Begin MSFlexGridLib.MSFlexGrid flx_data 
         Height          =   3495
         Left            =   120
         TabIndex        =   27
         Top             =   2040
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   6165
         _Version        =   393216
      End
      Begin VB.Label Label2 
         Caption         =   "Category"
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
         Index           =   7
         Left            =   7080
         TabIndex        =   49
         Top             =   120
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
         Index           =   6
         Left            =   5040
         TabIndex        =   26
         Top             =   120
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
         Index           =   5
         Left            =   360
         TabIndex        =   25
         Top             =   120
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
         Index           =   4
         Left            =   1680
         TabIndex        =   24
         Top             =   120
         Width           =   1095
      End
   End
   Begin VB.Frame Frame5 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   4560
      TabIndex        =   0
      Top             =   7320
      Width           =   2175
      Begin VB.CommandButton Refresh 
         BackColor       =   &H00C0E0FF&
         Caption         =   "&Refresh"
         Height          =   705
         Left            =   360
         MaskColor       =   &H000000FF&
         Picture         =   "frm_leave_chentry.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   46
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton exit 
         BackColor       =   &H00C0E0FF&
         Caption         =   "&Exit"
         Height          =   705
         Left            =   1200
         MaskColor       =   &H000000FF&
         Picture         =   "frm_leave_chentry.frx":066A
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   120
         Width           =   735
      End
   End
   Begin MSFlexGridLib.MSFlexGrid flx_dataold 
      Height          =   1455
      Left            =   6840
      TabIndex        =   31
      Top             =   7320
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   2566
      _Version        =   393216
   End
   Begin MSComCtl2.DTPicker dt_entdate 
      Height          =   375
      Left            =   10080
      TabIndex        =   50
      Top             =   840
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   61210625
      CurrentDate     =   42278
   End
   Begin VB.Label Label3 
      Caption         =   "Compensation Off Entries - for Employees"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   17
      Top             =   240
      Width           =   6255
   End
End
Attribute VB_Name = "frm_leave_chentry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim del_leave As Integer
Dim no As Integer
Dim rdate As Date

Private Sub cmb_leave_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub cmd_Assign_Click()
Dim i As Integer
   If opt_leave_half.Value = True Then
      If opt_an.Value = False And opt_fn.Value = False Then
         MsgBox ("Select Leave - FORE NOON / AFTER NOON ")
         Exit Sub
      End If
   End If
    
   For i = 1 To flx_data.Rows - 1
        If flx_data.TextMatrix(i, 3) <> "" Then
            If flx_data.TextMatrix(i, 2) = "" Then
                pst_qry = MsgBox("No Working hours on that day,Press YES-to continue  NO-to CANCEL", vbYesNo, "Confirmation")
                If pst_qry <> 6 Then
                    Exit Sub
                End If
            End If
            If (Option1.Value = True Or Option3.Value = True) And flx_data.TextMatrix(i, 5) = "F" And opt_leave_full.Value = True And Val(flx_data.TextMatrix(i, 2)) < 7.5 Then
                MsgBox "Working hours less than 8 Hours,Not eligible for fullday CH"
                Exit Sub
            ElseIf (Option1.Value = True Or Option3.Value = True) And flx_data.TextMatrix(i, 5) = "F" And opt_leave_full.Value = False And Val(flx_data.TextMatrix(i, 2)) < 7.5 Then
                MsgBox "Working hours less than 8 Hours,Not eligible for fullday CH"
                Exit Sub
            ElseIf (Option1.Value = True Or Option3.Value = True) And flx_data.TextMatrix(i, 5) = "H" And Val(flx_data.TextMatrix(i, 2)) < 4 Then
                
                MsgBox "Working hours less than 4 Hours,Not eligible for Halfday CH"
                Exit Sub
    ''        ElseIf Option2.Value = True And opt_leave_full.Value = True And Val(flx_data.TextMatrix(i, 2)) < 16 Then
            
            ElseIf Option2.Value = True And flx_data.TextMatrix(i, 5) = "F" And Val(flx_data.TextMatrix(i, 2)) < 15.55 Then
                MsgBox "Extra Working hours less than 8 Hours,Not eligible for fullday CH"
                Exit Sub
            ElseIf Option2.Value = True And opt_leave_half.Value = True And Val(flx_data.TextMatrix(i, 2)) < 4 Then
                MsgBox "Extra Working hours less than 4 Hours,Not eligible for Halfday CH"
                Exit Sub
            End If
''             If CDate(Format(flx_data.TextMatrix(i, 1), "MM/dd/yyyy")) > CDate(Format(flx_data.TextMatrix(i, 3), "MM/dd/yyyy")) Then
''            If CDate(flx_data.TextMatrix(i, 1)) > CDate(flx_data.TextMatrix(i, 3)) Then
''                MsgBox "CH Date should not less than worked date"
''                Exit Sub
''             End If
        End If
   Next
   
   
   For i = 1 To flx_data.Rows - 1
        If flx_data.TextMatrix(i, 9) = "" Then
        If flx_data.TextMatrix(i, 3) <> "" And flx_data.TextMatrix(i, 9) <> "" Then
            pst_qry = "select * from bio_emp_chleave where empch_fpcode =  " & Trim(txt_fpcode.Text) & "  And empch_ch_date = '" & Format(CDate(flx_data.TextMatrix(i, 3)), "MM/dd/yyyy") & "'"
            payrs.Open pst_qry, paydb, 1, 2
            If Not payrs.EOF Then
               MsgBox (" Already CH entries are made for " + flx_data.TextMatrix(i, 3) + " in entry No. " + Str(payrs("empch_no")) + " in the Date .." + Format(payrs("empch_entry_date"), "dd/MM/yyyy"))
               payrs.Close
               Exit Sub
            Else
              payrs.Close
            End If
        End If
        End If
   Next
   
   For i = 1 To flx_data.Rows - 1
        If flx_data.TextMatrix(i, 3) <> "" Then
            pst_qry = "select * from bio_empleave where emp_fpcode =  " & Trim(txt_fpcode.Text) & "  And emp_leave_date = '" & Format(CDate(flx_data.TextMatrix(i, 3)), "MM/dd/yyyy") & "'"
            payrs.Open pst_qry, paydb, 1, 2
            If Not payrs.EOF Then
               MsgBox (" Already Leave entries are made for " + flx_data.TextMatrix(i, 3))
               payrs.Close
               Exit Sub
            Else
              payrs.Close
            End If
        End If
   Next
   
   For i = 1 To flx_data.Rows - 1
        If flx_data.TextMatrix(i, 3) <> "" Then
            pst_qry = "select * from bio_device_shiftlogs where ds_fpcode =  " & Trim(txt_fpcode.Text) & "  And ds_date = '" & Format(CDate(flx_data.TextMatrix(i, 3)), "MM/dd/yyyy") & "'"
            payrs.Open pst_qry, paydb, 1, 2
            If Not payrs.EOF Then
               If payrs("ds_status") = "P" Then
                  MsgBox (" Employee Present ON " + flx_data.TextMatrix(i, 3) + ". You can't give CH")
                  payrs.Close
                  Exit Sub
               Else
                  payrs.Close
               End If
            Else
              payrs.Close
            End If
        End If
   Next
   
   
   
paydb.BeginTrans
On Error GoTo err_handler
    Dim ltype, chfrom, sql, chtime As String
    If opt_leave_full.Value = True Then
       ltype = "F"
    Else
       If opt_fn.Value = True Then
          ltype = "1"
       Else
          ltype = "2"
       End If
    End If
    If Option1.Value = True Then
        chfrom = "W"
    ElseIf Option3.Value = True Then
        chfrom = "H"
    Else
        chfrom = "E"
    End If
    
    chtime = "0"
    If opt_fn.Value = True Then chtime = "1"
    If opt_an.Value = True Then chtime = "2"
    
    
    For i = 1 To flx_dataold.Rows - 1
    Dim sdate As Date
    If Format(flx_dataold.TextMatrix(i, 1), "dd/MM/yyyy") <> "" Then
        sdate = Format(flx_dataold.TextMatrix(i, 1), "dd/MM/yyyy")
        sql = "delete from bio_emp_chleave where empch_fpcode = " & txt_fpcode.Text & "  and empch_worked_date  = '" & Format(sdate, "yyyy/MM/dd") & "' and empch_availedfrom = '" & chfrom & "'"
        paydb.Execute sql
    End If

    Next


    For i = 1 To flx_data.Rows - 1
    If flx_data.TextMatrix(i, 3) <> "" Then
        pst_qry = "select max(empch_no)+1 as endno from bio_emp_chleave"
        payrs.Open pst_qry, paydb, 1, 2
        no = 1
        If Not IsNull(payrs!endno) Then
            If Not payrs.EOF Then
                 no = payrs!endno
            End If
        End If
        payrs.Close
''            sql = "insert into bio_emp_chleave (empch_no , empch_fpcode, empch_worked_date,empch_worked_hours, empch_ch_date, emp_ch_period,empch_availedfrom) values (" & no & ", " & txt_fpcode.Text & ", '" & Format(flx_data.TextMatrix(i, 1), "yyyy-MM-dd") & "','" & Val(flx_data.TextMatrix(i, 2)) & "','" & Format(flx_data.TextMatrix(i, 3), "yyyy-MM-dd") & "','" & ltype & "','" & chfrom & "')"
        sql = "insert into bio_emp_chleave (empch_no ,empch_entry_date, empch_fpcode, empch_worked_date,empch_worked_hours, empch_ch_date, emp_ch_period,empch_availedfrom,empch_time) values (" & no & ",'" & Format(dt_entdate.Value, "MM/dd/yyyy") & "', " & txt_fpcode.Text & ", '" & Format(flx_data.TextMatrix(i, 1), "yyyy-MM-dd") & "','" & Val(flx_data.TextMatrix(i, 2)) & "','" & Format(flx_data.TextMatrix(i, 3), "yyyy-MM-dd") & "', '" & flx_data.TextMatrix(i, 5) & "' ,'" & flx_data.TextMatrix(i, 4) & "' , '" & chtime & "')"
            paydb.Execute sql
    End If
    Next

    paydb.CommitTrans
''    MsgBox "Record Saved", vbOKOnly + vbInformation, "Information"
    MsgBox "Record Saved in the Entry Number : " + Str(no) + " Entry Date : " + "'" & Format(dt_entdate.Value, "dd/MM/YYYY") & "'", vbOKOnly + vbInformation, "Information"""
    
    Refresh_Click
    Exit Sub
Exit Sub
err_handler:
        paydb.RollbackTrans
        chk = gen_Validation(Err.Number, Err.Description)

End Sub

Private Sub cmd_clear_Click()
          txt_fpcode.Text = ""
          txt_empname2.Text = ""
          txt_dept.Text = ""
End Sub

Private Sub cmd_delete_Click()
    del_leave = 1
    cmd_Assign.Enabled = False
    cmd_modify.Enabled = True
    flx_data.Enabled = True
End Sub

Private Sub cmd_filter_Click()

    Refresh_Click
    
    Dim payrs As New ADODB.Recordset
    
    If txt_empcode.Text <> "" Then
      sql = "select * from bio_empmas where bioemp_fpcode =  '" & txt_empcode.Text & "' and bioemp_status = 'Working' order by bioemp_dept"
    ElseIf txt_empname.Text <> "" Then
       sql = "select * from bio_empmas where bioemp_name like  '" & txt_empname.Text & "' and bioemp_status = 'Working' order by bioemp_dept"
    ElseIf lst_employee.Text <> "" Then
       sql = "select * from bio_empmas where bioemp_name =  '" & lst_employee.Text & "' and bioemp_status = 'Working' order by bioemp_dept"
    Else
       MsgBox ("Employee not selected...")
       Exit Sub
    End If
    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
    While Not payrs.EOF()
          txt_fpcode.Text = payrs!bioemp_fpcode
          txt_empname2.Text = payrs!bioemp_name
          txt_dept.Text = payrs!bioemp_dept
          txt_cat.Text = payrs!bioemp_team
          payrs.MoveNext
    Wend
    payrs.Close
''    If Option1.Value = True Then
''        pst_qry = " select empch_worked_date as wdate,empch_worked_hours as whours,empch_ch_date as chdate,empch_availedfrom as chfrom,emp_ch_period as day from bio_emp_chleave where empch_fpcode='" & txt_fpcode.Text & "' and empch_availedfrom in ('W') and empch_worked_date  between '" & Format(st_date, "MM/dd/yyyy") & "' and '" & Format(end_date, "MM/dd/yyyy") & "' union all select ds_date as wdate,ds_sft_hrs as whours,null as chdate,'' as chfrom,'' as day from bio_device_shiftlogs  where  ds_fpcode =  '" & txt_fpcode.Text & "' and ds_status in ('WOP','WO½P','WOP(OD)','WOHP','WOHPE') and  ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and '" & Format(end_date, "MM/dd/yyyy") & "' and ds_date not in (select empch_worked_date from bio_emp_chleave where empch_fpcode='" & txt_fpcode.Text & "' and empch_availedfrom in ('W'))order by empch_worked_date  "
''    ElseIf Option3.Value = True Then
''        pst_qry = " select empch_worked_date as wdate,empch_worked_hours as whours,empch_ch_date as chdate,empch_availedfrom as chfrom,emp_ch_period as day  from bio_emp_chleave where empch_fpcode='" & txt_fpcode.Text & "' and empch_availedfrom in ('H') and empch_worked_date  between '" & Format(st_date, "MM/dd/yyyy") & "' and '" & Format(end_date, "MM/dd/yyyy") & "' union all select ds_date as wdate,ds_sft_hrs as whours,null as chdate,'' as chfrom,null as day from bio_device_shiftlogs  where  ds_fpcode =  '" & txt_fpcode.Text & "' and ds_status in ('HP','HPE','½HP') and  ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and '" & Format(end_date, "MM/dd/yyyy") & "' and ds_date not in (select empch_worked_date from bio_emp_chleave where empch_fpcode='" & txt_fpcode.Text & "' and empch_availedfrom in ('H')) order by empch_worked_date "
''    Else
''         pst_qry = "select empch_worked_date as wdate,empch_worked_hours as whours,empch_ch_date as chdate,empch_availedfrom as chfrom,emp_ch_period as day from bio_emp_chleave where empch_fpcode= '" & txt_fpcode.Text & "' and empch_availedfrom in ('E') and empch_worked_date  between '" & Format(st_date, "MM/dd/yyyy") & "' and '" & Format(end_date, "MM/dd/yyyy") & "'  union all  select ds_date as wdate,ds_sft_hrs as whours,null as chdate,'' as chfrom,'' as day  from bio_device_shiftlogs  where  ds_fpcode =  '" & txt_fpcode.Text & "' and  ds_sft_hrs >=12  and  ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and '" & Format(end_date, "MM/dd/yyyy") & "'  and ds_date not in (select empch_worked_date from bio_emp_chleave where empch_fpcode= '" & txt_fpcode.Text & "' and empch_availedfrom in ('E')) order by empch_worked_date"
''    End If

    If Option1.Value = True Then
        pst_qry = " select empch_worked_date as wdate,empch_worked_hours as whours,empch_ch_date as chdate,empch_availedfrom as chfrom,emp_ch_period as day ,empch_no,empch_entry_date  from bio_emp_chleave where empch_fpcode='" & txt_fpcode.Text & "' and empch_availedfrom in ('W') and empch_worked_date  between '" & Format(st_date, "MM/dd/yyyy") & "' and '" & Format(end_date, "MM/dd/yyyy") & "' union all select ds_date as wdate,ds_sft_hrs as whours,null as chdate,'' as chfrom,'' as day ,0 as empch_no,null as  empch_entry_date  from bio_device_shiftlogs  where  ds_fpcode =  '" & txt_fpcode.Text & "' and ds_status in ('WOP','WO½P','WOP(OD)','WOHP','WOHPE') and  ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and '" & Format(end_date, "MM/dd/yyyy") & "' and ds_date not in (select empch_worked_date from bio_emp_chleave where empch_fpcode='" & txt_fpcode.Text & "' and empch_availedfrom in ('W'))order by empch_worked_date  "
    ElseIf Option3.Value = True Then
        pst_qry = " select empch_worked_date as wdate,empch_worked_hours as whours,empch_ch_date as chdate,empch_availedfrom as chfrom,emp_ch_period as day ,empch_no,empch_entry_date   from bio_emp_chleave where empch_fpcode='" & txt_fpcode.Text & "' and empch_availedfrom in ('H') and empch_worked_date  between '" & Format(st_date, "MM/dd/yyyy") & "' and '" & Format(end_date, "MM/dd/yyyy") & "' union all select ds_date as wdate,ds_sft_hrs as whours,null as chdate,'' as chfrom,null as day  ,0 as empch_no,null as  empch_entry_date  from bio_device_shiftlogs  where  ds_fpcode =  '" & txt_fpcode.Text & "' and ds_status in ('HP','HPE','½HP') and  ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and '" & Format(end_date, "MM/dd/yyyy") & "' and ds_date not in (select empch_worked_date from bio_emp_chleave where empch_fpcode='" & txt_fpcode.Text & "' and empch_availedfrom in ('H')) order by empch_worked_date "
    Else
         pst_qry = "select empch_worked_date as wdate,empch_worked_hours as whours,empch_ch_date as chdate,empch_availedfrom as chfrom,emp_ch_period as day ,empch_no,empch_entry_date from bio_emp_chleave where empch_fpcode= '" & txt_fpcode.Text & "' and empch_availedfrom in ('E') and empch_worked_date  between '" & Format(st_date, "MM/dd/yyyy") & "' and '" & Format(end_date, "MM/dd/yyyy") & "'  union all  select ds_date as wdate,ds_sft_hrs as whours,null as chdate,'' as chfrom,'' as day  ,0 as empch_no,null as  empch_entry_date   from bio_device_shiftlogs  where  ds_fpcode =  '" & txt_fpcode.Text & "' and  ds_sft_hrs >=12  and  ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and '" & Format(end_date, "MM/dd/yyyy") & "'  and ds_date not in (select empch_worked_date from bio_emp_chleave where empch_fpcode= '" & txt_fpcode.Text & "' and empch_availedfrom in ('E')) order by empch_worked_date"
    End If

    payrs.Open pst_qry, paydb, adOpenDynamic, adLockOptimistic
    i = 1
    While Not payrs.EOF
        flx_data.TextMatrix(i, 0) = i
        If Option1.Value = True Then
            flx_data.TextMatrix(i, 1) = Format(payrs!wdate, "dd/MM/yyyy")
            flx_data.TextMatrix(i, 2) = Format(payrs!whours, "00.00")
            flx_data.TextMatrix(i, 4) = payrs!chfrom
            flx_data.TextMatrix(i, 5) = payrs!Day
  '          flx_data.TextMatrix(i, 3) = Format(payrs!chdate, "dd/MM/yyyy")
        ElseIf Option3.Value = True Then
            flx_data.TextMatrix(i, 1) = Format(payrs!wdate, "dd/MM/yyyy")
            flx_data.TextMatrix(i, 2) = Format(payrs!whours, "00.00")
            flx_data.TextMatrix(i, 4) = "H"
            flx_data.TextMatrix(i, 5) = IIf(IsNull(payrs!Day), "", payrs!Day)
   ''         flx_data.TextMatrix(i, 3) = Format(payrs!chdate, "dd/MM/yyyy")
        Else
'            flx_data.TextMatrix(i, 1) = Format(payrs!ds_date, "dd/MM/yyyy")
 '           flx_data.TextMatrix(i, 2) = Format(payrs!ds_sft_hrs, "00.00")
            flx_data.TextMatrix(i, 1) = Format(payrs!wdate, "dd/MM/yyyy")
            flx_data.TextMatrix(i, 2) = Format(payrs!whours, "00.00")
            flx_data.TextMatrix(i, 4) = "E"
            flx_data.TextMatrix(i, 5) = IIf(IsNull(payrs!Day), "", payrs!Day)
        End If
        If Option1.Value = True Or Option3.Value = True Then
           flx_data.TextMatrix(i, 7) = Format(Int(payrs!whours), "00.00")
        Else
           flx_data.TextMatrix(i, 7) = Format(Int(payrs!whours - 8), "00.00")
        End If
        flx_data.TextMatrix(i, 3) = Format(payrs!chdate, "dd/MM/yyyy")
        
        flx_data.TextMatrix(i, 9) = payrs!empch_no
        flx_data.TextMatrix(i, 10) = Format(payrs!empch_entry_date, "dd/MM/yyyy")
             
        
        flx_data.Rows = flx_data.Rows + 1
        flx_dataold.TextMatrix(i, 0) = i
        flx_dataold.TextMatrix(i, 1) = Format(payrs!wdate, "dd/MM/yyyy")
        flx_dataold.TextMatrix(i, 2) = Format(payrs!whours, "00.00")
        flx_dataold.TextMatrix(i, 3) = Format(payrs!chdate, "dd/MM/yyyy")
        flx_dataold.TextMatrix(i, 4) = Format(payrs!chfrom, "00.00")
        flx_dataold.TextMatrix(i, 5) = IIf(IsNull(payrs!Day), "", payrs!Day)
        flx_dataold.TextMatrix(i, 7) = Format(Int(payrs!whours - 8), "00.00")
        flx_dataold.TextMatrix(i, 9) = payrs!empch_no
        flx_dataold.TextMatrix(i, 10) = Format(payrs!empch_entry_date, "dd/MM/yyyy")
        
        payrs.MoveNext
        flx_dataold.Rows = flx_dataold.Rows + 1
        
        i = i + 1
    Wend
    payrs.Close
''for updating permissions
     Dim time1, ftime, etime As Double
     Dim workedtotmins, workedhrs, workedmins, tothrs As Double
    For i = 1 To flx_data.Rows - 1
        pst_qry = "select * from bio_emp_permissions where empp_fpcode =  '" & txt_fpcode.Text & "' and  empp_date = '" & Format(flx_data.TextMatrix(i, 1), "MM/dd/yyyy") & "'"
        payrs.Open pst_qry, paydb, adOpenDynamic, adLockOptimistic
        While Not payrs.EOF
              If flx_data.TextMatrix(i, 1) = Format(payrs!empp_date, "dd/MM/yyyy") Then
                 time1 = Int(Val(flx_data.TextMatrix(i, 2))) * 60 + (Val(flx_data.TextMatrix(i, 2)) - Int(Val(flx_data.TextMatrix(i, 2)))) * 100
                 ftime = Int(Val(payrs!empp_fromtime)) * 60 + (Val(payrs!empp_fromtime) - Int(Val(payrs!empp_fromtime))) * 100
                 etime = Int(Val(payrs!empp_totime)) * 60 + (Val(payrs!empp_totime) - Int(Val(payrs!empp_totime))) * 100
                 workedtotmins = time1 + (etime - ftime)
                 workedhrs = Int(workedtotmins / 60)
                 workedmins = workedtotmins - (workedhrs * 60)
                 tothrs = workedhrs + (workedmins / 100)
                 flx_data.TextMatrix(i, 2) = Format(tothrs, "0.00")
                 flx_data.TextMatrix(i, 7) = Format(Int(tothrs - 8), "00.00")
              End If
              payrs.MoveNext
     
        Wend
        payrs.Close
    Next
      
      
    pst_qry = " select w_date ,w_accepted_hrs from bio_worker_daily_pihrs where w_emp_fpcode='" & txt_fpcode.Text & "' and w_date  between '" & Format(st_date, "MM/dd/yyyy") & "' and '" & Format(end_date, "MM/dd/yyyy") & "' order by w_date"
    payrs.Open pst_qry, paydb, adOpenDynamic, adLockOptimistic
    While Not payrs.EOF
      For i = 1 To flx_data.Rows - 1
         If flx_data.TextMatrix(i, 1) = Format(payrs!w_date, "dd/MM/yyyy") Then
            flx_data.TextMatrix(i, 8) = Format(payrs!w_accepted_hrs, "0.00")
         End If
      Next
      payrs.MoveNext
    Wend
    payrs.Close
    
    
    
''        pst_qry = "select * from bio_emp_chleave  where empch_fpcode =  '" & txt_fpcode.Text & "' and empch_availedfrom in ('W')  and  empch_worked_date between '" & Format(st_date, "MM/dd/yyyy") & "' and '" & Format(end_date, "MM/dd/yyyy") & "'"
''        payrs.Open pst_qry, paydb, adOpenDynamic, adLockOptimistic
''       i = 1
''    While Not payrs.EOF
''    For i = 1 To flx_data.Rows - 1
''
''        If flx_data.TextMatrix(i, 1) = Format(payrs!empch_worked_date, "dd/MM/yyyy") Then
''             flx_data.TextMatrix(i, 3) = Format(payrs!empch_ch_date, "dd/MM/yyyy")
''        End If
''    Next
''        payrs.MoveNext
''''        flx_data.Rows = flx_data.Rows + 1
''    i = i + 1
''
''''    Next
''    Wend
''    payrs.Close
''''    flx_data.Enabled = False
 

    
End Sub

Private Sub cmd_modify_Click()
Dim i As Integer
paydb.BeginTrans
On Error GoTo err_handler
    Dim ltype, chfrom, sql As String
    If opt_leave_full.Value = True Then
       ltype = "F"
    Else
       If opt_fn.Value = True Then
          ltype = "1"
       Else
          ltype = "2"
       End If
    End If
    If Option1.Value = True Then
        chfrom = "W"
    ElseIf Option3.Value = True Then
        chfrom = "H"
    Else
        chfrom = "E"
    End If
    
    For i = 1 To flx_dataold.Rows - 1
    Dim sdate As Date
    If Format(flx_dataold.TextMatrix(i, 1), "dd/MM/yyyy") <> "" Then
        sdate = Format(flx_dataold.TextMatrix(i, 1), "dd/MM/yyyy")
        sql = "delete from bio_emp_chleave where empch_fpcode = " & txt_fpcode.Text & "  and empch_worked_date  = '" & Format(sdate, "yyyy/MM/dd") & "'"
        paydb.Execute sql
    End If

    Next


    For i = 1 To flx_data.Rows - 1
    If flx_data.TextMatrix(i, 3) <> "" Then
        pst_qry = "select max(empch_no)+1 as endno from bio_emp_chleave"
        payrs.Open pst_qry, paydb, 1, 2
        no = 1
        If Not IsNull(payrs!endno) Then
            If Not payrs.EOF Then
                 no = payrs!endno
            End If
        End If
        payrs.Close
        sql = "insert into bio_emp_chleave (empch_no , empch_fpcode, empch_worked_date,empch_worked_hours, empch_ch_date, emp_ch_period,empch_availedfrom) values (" & no & ", " & txt_fpcode.Text & ", '" & Format(flx_data.TextMatrix(i, 1), "yyyy-MM-dd") & "','" & Val(flx_data.TextMatrix(i, 2)) & "','" & Format(flx_data.TextMatrix(i, 3), "yyyy-MM-dd") & "','" & ltype & "','" & chfrom & "')"
        paydb.Execute sql
    End If
    Next

    paydb.CommitTrans
    MsgBox "Record Saved", vbOKOnly + vbInformation, "Information"
    
    Refresh_Click
    Exit Sub
Exit Sub
err_handler:
        paydb.RollbackTrans
        chk = gen_Validation(Err.Number, Err.Description)

End Sub

Private Sub exit_Click()
     Unload Me
End Sub

Private Sub flx_data_Click()
On Error GoTo err_handler
    flx_data_validation
    Exit Sub
err_handler:
    chk = gen_Validation(Err.Number, Err.Description)
    If chk = 1 Then Resume
End Sub

Private Sub flx_data_DblClick()
   If del_leave = 0 Then Exit Sub
   flex_edit_row = 0
   Dim fin_selrow As Integer
   Dim chdate, duedate, mstartdate As Date
   Dim pst_ans As String
   fin_selrow = flx_data.Row

   timchk = 0
   With flx_data
       pst_ans = MsgBox("Press YES-to DELETE  NO-to CANCEL", vbYesNo, "Confirmation")
       If pst_ans = 6 Then
               If .Rows < 2 Then
                  MsgBox "No rows to remove"
               Else
                  If Val(flx_data.TextMatrix(.Row, 0)) > 0 Then
                     chdate = CDate(flx_data.TextMatrix(.Row, 3))
                     duedate = chdate + 5
                     mstartdate = dt_entdate.Value - Day(dt_entdate.Value) + 1
                  
                     If Month(dt_entdate.Value) <> Format(chdate, "MM/dd/YYYY") Then
''                        If Format(dt_entdate.Value, "MM/dd/YYYY") > Format(duedate, "MM/dd/YYYY") Then
''                           MsgBox ("You Can't Delete .. Last Month CH Entries...")
''                           Exit Sub
''                         End If
                      End If
                  
                      flx_data.RemoveItem fin_selrow
                      .Row = flx_data.Rows - 1
                  End If
              End If
        End If
    End With
End Sub

Private Sub flx_data_EnterCell()
On Error GoTo err_handler
    flx_data_validation
    Exit Sub
err_handler:
    chk = gen_Validation(Err.Number, Err.Description)
    If chk = 1 Then Resume
End Sub



''Private Sub flx_edit_DblClick()
''   flex_edit_row = 0
''   Dim fin_selrow As Integer
''   Dim pst_ans As String
''   fin_selrow = flx_edit.Row
''
''   With flx_edit
''       pst_ans = MsgBox("Press YES-to Modify  NO-to Delete", vbYesNo, "Confirmation")
''       If pst_ans = 6 Then
''
''        Else
''        If .Rows < 2 Then
''                  MsgBox "No rows to remove"
''               Else
''                  If Val(flx_edit.TextMatrix(.Row, 0)) > 0 Then
''                     flx_edit.RemoveItem fin_selrow
''                  End If
''                  .Row = flx_edit.Rows - 1
''               End If
''        End If
''   End With
''End Sub

Private Sub Form_Load()
    del_leave = 0
    dt_date.Value = Now
    dt_entdate.Value = Now
    fillgrid
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
''        .AddItem finyear + 2000
''        .AddItem "2012"
''        .AddItem "2013"
''        .AddItem "2014"
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



End Sub

Private Sub lst_dept_Click()
    lst_employee.Clear
    sql = "Select * from  bio_empmas where bioemp_status = 'Working'  and bioemp_dept = '" & lst_dept.Text & "' order by bioemp_name "
    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
    While Not payrs.EOF()
        lst_employee.AddItem payrs("bioemp_name")
        lst_employee.ItemData(lst_employee.NewIndex) = payrs("bioemp_fpcode")
        payrs.MoveNext
    Wend
    payrs.Close
End Sub

Private Sub opt_leave_full_Click()
    frame_leavetype.Visible = False
 
End Sub
''Public Sub leave_details()
'' fillgrid
''    Dim payrs As New ADODB.Recordset
''    pst_qry = "select * from bio_emp_chleave  a ,emp_mas b  where  a.empch_fpcode = b.emp_fpcode and b.emp_fpcode =  '" & txt_fpcode.Text & "' and  empch_ch_date between '" & Format(st_date, "MM/dd/yyyy") & "' and '" & Format(end_date, "MM/dd/yyyy") & "'"
''    payrs.Open pst_qry, paydb, adOpenDynamic, adLockOptimistic
''    i = 1
''    While Not payrs.EOF
''        flx_data.TextMatrix(i, 0) = i
''        flx_data.TextMatrix(i, 1) = Format(payrs!empch_worked_date, "dd/MM/yyyy")
''        flx_data.TextMatrix(i, 2) = Format(payrs!empch_ch_date, "dd/MM/yyyy")
''        flx_data.TextMatrix(i, 3) = payrs!empch_availedfrom
''        flx_data.TextMatrix(i, 4) = payrs!emp_ch_period
''        flx_data.Rows = flx_data.Rows + 1
''
''        flx_dataold.TextMatrix(i, 0) = i
''        flx_dataold.TextMatrix(i, 1) = Format(payrs!empch_worked_date, "dd/MM/yyyy")
''        flx_dataold.TextMatrix(i, 2) = Format(payrs!empch_ch_date, "dd/MM/yyyy")
''        flx_dataold.TextMatrix(i, 3) = payrs!empch_availedfrom
''        flx_dataold.TextMatrix(i, 4) = payrs!emp_ch_period
''        flx_dataold.Rows = flx_dataold.Rows + 1
''        payrs.MoveNext
''
''        i = i + 1
''    Wend
''    payrs.Close
''  flx_data.Enabled = False
''End Sub

Public Sub fillgrid()


   With flx_data
     .Redraw = False
     .Clear
     .Rows = 2
     .Cols = 11
     .TextMatrix(0, 0) = "S.No"
     .TextMatrix(0, 1) = "Worked Date"
     .TextMatrix(0, 2) = "Worked Hrs"
     .TextMatrix(0, 3) = "CH Date"
     .TextMatrix(0, 4) = "CHfrom"
     .TextMatrix(0, 5) = "Full/Half"
     .TextMatrix(0, 6) = "Noon"
     .TextMatrix(0, 7) = "EX.HRS"
     .TextMatrix(0, 8) = "OT AVAILED"
     .TextMatrix(0, 9) = "ENT.NO"
     .TextMatrix(0, 10) = "ENT.DATE"
     
     .ColWidth(0) = 500
     .ColWidth(1) = 2000
     .ColWidth(2) = 1000
     .ColWidth(3) = 1300
     .ColWidth(4) = 1000
     .ColWidth(5) = 1000
     .ColWidth(6) = 1000
     .ColWidth(7) = 1000
     .ColWidth(8) = 1000
     .ColWidth(9) = 1000
     .ColWidth(10) = 1000
     
     .Redraw = True
   End With
   With flx_dataold
     .Redraw = False
     .Clear
     .Rows = 2
     .Cols = 11
     .TextMatrix(0, 0) = "S.No"
     .TextMatrix(0, 1) = "Worked Date(WO/H)"
     .TextMatrix(0, 2) = "Worked Hrs"
     .TextMatrix(0, 3) = "CH Date"
     .TextMatrix(0, 4) = "CHfrom"
     .TextMatrix(0, 5) = "Full/Half"
     .TextMatrix(0, 6) = "Noon"
     .TextMatrix(0, 7) = "EX.HRS"
     .TextMatrix(0, 8) = "OT AVAILED"
     .TextMatrix(0, 9) = "ENT.NO"
     .TextMatrix(0, 10) = "ENT.DATE"
     
     .ColWidth(0) = 500
     .ColWidth(1) = 2000
     .ColWidth(2) = 1000
     .ColWidth(3) = 1300
     .ColWidth(4) = 1000
     .ColWidth(5) = 1000
     .ColWidth(6) = 1000
     .ColWidth(7) = 1000
     .ColWidth(8) = 1000
     .ColWidth(9) = 1000
     .ColWidth(10) = 1000

     .Redraw = True
   End With
End Sub
Private Sub cmb_month_Click()
    find_dates
End Sub

Private Sub cmb_year_Click()
   find_dates
''    Label5.Caption = cmb_month.Text & "-" & cmb_year.Text
End Sub
Public Sub find_dates()
    If cmb_month.ListIndex = -1 Then Exit Sub
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
End Sub


Sub flx_data_validation()
On Error GoTo err_handler
    If txt_cat.Text <> "STAFF" Then
        If Val(flx_data.TextMatrix(flx_data.Row, 7)) - Val(flx_data.TextMatrix(flx_data.Row, 8)) < 2 Then
           MsgBox ("Already PI hrs availed...")
           Exit Sub
        End If
    End If
    Select Case flx_data.Col
        Case 3
            If flx_data.TextMatrix(flx_data.Row, 1) <> "" Then
                dt_date.Width = flx_data.ColWidth(3)
                dt_date.Top = flx_data.Top + flx_data.CellTop
                dt_date.Left = flx_data.Left + flx_data.CellLeft
                If flx_data.TextMatrix(flx_data.Row, 3) <> "" Then
                    dt_date.Value = CDate(Format$(flx_data.TextMatrix(flx_data.Row, 3), "dd/mm/yyyy"))
                End If
                dt_date.Visible = True
                dt_date.SetFocus
                
            End If
    End Select
    Exit Sub
err_handler:
    chk = gen_Validation(Err.Number, Err.Description)
    'If chk = 1 Then Resume
End Sub


Private Sub flx_data_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo err_handler
    If KeyCode = 46 And flx_data.Col = 3 Then
''        flx_data.TextMatrix(flx_data.Row, 2) = ""
    End If
    Exit Sub
err_handler:
    chk = gen_Validation(Err.Number, Err.Description)
    If chk = 1 Then Resume
End Sub

Private Sub flx_data_KeyPress(KeyAscii As Integer)
On Error GoTo err_handler
    flx_data_validation
''    Select Case flx_data.Col
''        Case 7
''            If Len(Trim(flx_data.TextMatrix(flx_data.Row, 7))) >= 15 Or KeyAscii = 39 Then
''                flx_data.TextMatrix(flx_data.Row, 7) = flx_data.TextMatrix(flx_data.Row, 7)
''            ElseIf KeyAscii = 8 Then
''                If Trim(flx_data.TextMatrix(flx_data.Row, 7)) <> "" Then
''                    flx_data.TextMatrix(flx_data.Row, 7) = Mid(flx_data.TextMatrix(flx_data.Row, 7), 1, Len(flx_data.TextMatrix(flx_data.Row, 7)) - 1)
''                End If
''            Else
''                flx_data.TextMatrix(flx_data.Row, 7) = flx_data.TextMatrix(flx_data.Row, 7) + Chr(KeyAscii)
''            End If
''            If Trim(flx_data.TextMatrix(flx_data.Row, 7)) = "" Then
''                flx_data.TextMatrix(flx_data.Row, 8) = ""
''            End If
''    End Select
    Exit Sub
err_handler:
    chk = gen_Validation(Err.Number, Err.Description)
    If chk = 1 Then Resume
End Sub

Private Sub flx_data_LostFocus()
On Error GoTo err_handler
''    If flx_data.Col = 7 Then
''        flx_data.Col = 8
''        flx_data.SetFocus
''    End If
    Exit Sub
err_handler:
    chk = gen_Validation(Err.Number, Err.Description)
    If chk = 1 Then Resume
End Sub



''Private Sub dt_date_GotFocus()
''On Error GoTo err_handler
''    If flx_data.TextMatrix(flx_data.Row, 2) <> "" And flx_data.Row > 0 Then
''        dt_date.Value = CDate(flx_data.TextMatrix(flx_data.Row, 2))
''    End If
''    Exit Sub
''err_handler:
''    chk = gen_Validation(Err.Number, Err.Description)
''    If chk = 1 Then Resume
''End Sub

Private Sub dt_date_LostFocus()
''    MsgBox (dt_date)
    
    
    refresh_grid
End Sub

Public Sub refresh_grid()
    On Error GoTo err_handler
        If flx_data.TextMatrix(flx_data.Row, 1) <> "" Then
        
            If opt_leave_half.Value = True And (opt_an.Value = False And opt_fn.Value = False) Then
               MsgBox ("Select Fore / After Noon option and continue...")
               Exit Sub
            End If
        
            flx_data.TextMatrix(flx_data.Row, 3) = Format(dt_date.Value, "dd/MM/yyyy")
            If Option1.Value = True Then
               flx_data.TextMatrix(flx_data.Row, 4) = "W"
               flx_data.TextMatrix(flx_data.Row, 4) = "W"
            ElseIf Option3.Value = True Then
               flx_data.TextMatrix(flx_data.Row, 4) = "H"
            ElseIf Option2.Value = True Then
               flx_data.TextMatrix(flx_data.Row, 4) = "E"
            End If
''            If opt_leave_full.Value = True Then
''               ltype = "F"
''            Else
''               If opt_fn.Value = True Then
''                  ltype = "1"
''               Else
''                  ltype = "2"
''               End If
''            End If
''            flx_data.TextMatrix(flx_data.Row, 6) = ltype
            
            If opt_leave_full.Value = True Then
               ltype = "F"
               flx_data.TextMatrix(flx_data.Row, 5) = "F"
            Else
               flx_data.TextMatrix(flx_data.Row, 5) = "H"
               If opt_fn.Value = True Then
                  ltype = "1"
                  flx_data.TextMatrix(flx_data.Row, 6) = 1
               Else
                  ltype = "2"
                  flx_data.TextMatrix(flx_data.Row, 6) = 2
               End If
            End If
            ''flx_data.TextMatrix(flx_data.Row, 6) = ltype
            
            
            dt_date.Visible = False
        End If
        Exit Sub
err_handler:
        chk = gen_Validation(Err.Number, Err.Description)
        If chk = 1 Then Resume

End Sub

Private Sub opt_leave_half_Click()
   frame_leavetype.Visible = True
 End Sub

Private Sub Option1_Click()
   fillgrid
End Sub

Private Sub Option2_Click()
    Dim firsttime, secondtime, shrs, smins As Double
    fillgrid
    ''pst_qry = "select * from bio_device_shiftlogs  where  ds_fpcode =  '" & txt_fpcode.Text & "' and ds_sft_hrs+ds_sft_hrs2 >=12 and  ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and '" & Format(end_date, "MM/dd/yyyy") & "' and ds_date not in (select empch_worked_date from bio_emp_chleave where empch_fpcode='" & txt_fpcode.Text & "') order by ds_date"
    pst_qry = "select * from bio_device_shiftlogs  where  ds_fpcode =  '" & txt_fpcode.Text & "' and ds_sft_hrs >=12 and  ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and '" & Format(end_date, "MM/dd/yyyy") & "' and ds_date not in (select empch_worked_date from bio_emp_chleave where empch_fpcode='" & txt_fpcode.Text & "') order by ds_date"
    payrs.Open pst_qry, paydb, adOpenDynamic, adLockOptimistic
    i = 1
    While Not payrs.EOF
        firsttime = 0
        secondtime = 0
        shrs = 0
        smins = 0
        flx_data.TextMatrix(i, 0) = i
        flx_data.TextMatrix(i, 1) = Format(payrs!ds_date, "dd/MM/yyyy")
        If payrs!ds_sft_hrs2 > 0 Then
           firsttime = Int(payrs!ds_sft_hrs) + Int(payrs!ds_sft_hrs2)
           secondtime = ((payrs!ds_sft_hrs - Int(payrs!ds_sft_hrs)) + (payrs!ds_sft_hrs2 - Int(payrs!ds_sft_hrs2))) * 100
           If secondtime > 59 Then
              shrs = Int(secondtime / 60)
              smins = (secondtime - (shrs * 60)) / 100
           Else
              smins = secondtime / 100
           End If
           firsttime = firsttime + shrs + smins
           flx_data.TextMatrix(i, 2) = Format(firsttime, "00.00")
           flx_data.TextMatrix(i, 7) = Format(Int(firsttime - 8), "00.00")
        Else
           flx_data.TextMatrix(i, 2) = Format(payrs!ds_sft_hrs, "00.00")
           flx_data.TextMatrix(i, 7) = Format(Int(payrs!ds_sft_hrs - 8), "00.00")
        End If
        
        flx_data.Rows = flx_data.Rows + 1
        
        flx_dataold.TextMatrix(i, 0) = i
        flx_dataold.TextMatrix(i, 1) = Format(payrs!ds_date, "dd/MM/yyyy")
        payrs.MoveNext
        flx_dataold.Rows = flx_dataold.Rows + 1
        
        i = i + 1
    Wend
    payrs.Close
    
    pst_qry = " select w_date ,w_accepted_hrs from bio_worker_daily_pihrs where w_emp_fpcode='" & txt_fpcode.Text & "' and w_date  between '" & Format(st_date, "MM/dd/yyyy") & "' and '" & Format(end_date, "MM/dd/yyyy") & "' order by w_date"
    payrs.Open pst_qry, paydb, adOpenDynamic, adLockOptimistic
    While Not payrs.EOF
      For i = 1 To flx_data.Rows - 1
         If flx_data.TextMatrix(i, 1) = Format(payrs!w_date, "dd/MM/yyyy") Then
            flx_data.TextMatrix(i, 8) = Format(payrs!w_accepted_hrs, "0.00")
         End If
      Next
      payrs.MoveNext
    Wend
    payrs.Close
    
    
    
End Sub

Private Sub Option3_Click()
fillgrid
End Sub

Private Sub Refresh_Click()
    del_leave = 0
    cmd_Assign.Enabled = True
    cmd_modify.Enabled = False
    txt_fpcode.Text = ""
    txt_empname2.Text = ""
    txt_dept.Text = ""
    fillgrid
    flx_data.Enabled = True
End Sub

