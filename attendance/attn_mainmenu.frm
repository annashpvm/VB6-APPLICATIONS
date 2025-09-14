VERSION 5.00
Begin VB.MDIForm attn_mainmenu 
   BackColor       =   &H8000000C&
   Caption         =   "ATTENDANCE SYSTEM - MAIN MENU"
   ClientHeight    =   6855
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   11535
   LinkTopic       =   "ATTENDANCE SYSTEM - MAIN MENU"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu MASTER 
      Caption         =   "MASTER"
      Begin VB.Menu WEEK_OFF_MAS 
         Caption         =   "WEEKLY OFF MASTER"
      End
      Begin VB.Menu LEAVE_MASTER 
         Caption         =   "LEAVE MASTER"
      End
   End
   Begin VB.Menu UPD_SCR 
      Caption         =   "UPDATION SCREEN"
      Begin VB.Menu data_dpm 
         Caption         =   "Data Updation "
         Visible         =   0   'False
      End
      Begin VB.Menu data_new 
         Caption         =   "Data Updation NEW"
      End
   End
   Begin VB.Menu transactions 
      Caption         =   "TRANSACTIONS"
      Begin VB.Menu c_shift 
         Caption         =   "C SHIFT ENTRY "
      End
   End
   Begin VB.Menu REPORTS 
      Caption         =   "REPORTS"
      Begin VB.Menu attn_report 
         Caption         =   "Attendance Report"
      End
      Begin VB.Menu late_comers 
         Caption         =   "Late Comers List"
      End
      Begin VB.Menu IN_OUT_REP 
         Caption         =   "IN & OUT TIME REPORT"
      End
   End
   Begin VB.Menu WINDOW 
      Caption         =   "WINDOW"
      WindowList      =   -1  'True
   End
   Begin VB.Menu EXIT 
      Caption         =   "EXIT"
   End
End
Attribute VB_Name = "attn_mainmenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub attn_report_Click()
     rep_attn_daily.Show
End Sub

Private Sub c_shift_Click()
     shift_entry.Show
End Sub

Private Sub data_dpm_Click()
      attn_upd.Show
End Sub

Private Sub data_new_Click()
      attn_upd_new.Show
'      attn_upd.Show
End Sub

Private Sub EXIT_Click()
     Unload Me
End Sub


Private Sub IN_OUT_REP_Click()
     rep_attn_inouttime.Show
End Sub

Private Sub late_comers_Click()
     rep_attn_latecomers.Show
End Sub

Private Sub WEEK_OFF_MAS_Click()
     Weeklyoffmaster.Show
End Sub
