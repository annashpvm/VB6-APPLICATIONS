VERSION 5.00
Begin VB.MDIForm frm_mainmenu 
   BackColor       =   &H8000000C&
   Caption         =   "WEIGH BRIDGE SYSTEM"
   ClientHeight    =   8790
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   16710
   Icon            =   "frm_mainmenu.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu mnu_masters 
      Caption         =   "Masters"
      Begin VB.Menu mnu_item_group 
         Caption         =   "Materil Group Master"
      End
      Begin VB.Menu mnu_Material 
         Caption         =   "Material Master"
      End
      Begin VB.Menu mnu_Party 
         Caption         =   "Party Master"
      End
      Begin VB.Menu mnu 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_port 
         Caption         =   "Port Configuration"
      End
   End
   Begin VB.Menu mnu_weigh_entry 
      Caption         =   "Weigh Bridge Entry"
      Begin VB.Menu mnu_first 
         Caption         =   "First Transaction"
      End
      Begin VB.Menu mnu_second 
         Caption         =   "Second Transaction"
      End
      Begin VB.Menu mnu_import 
         Caption         =   "IMPORT"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnu_reports 
      Caption         =   "Reports"
      Begin VB.Menu mu_duplicate_ticket 
         Caption         =   "Duplicate Ticket"
      End
      Begin VB.Menu mnu_weightment_Details 
         Caption         =   "Weightment Details"
      End
   End
   Begin VB.Menu mnu_exit 
      Caption         =   "Exit"
   End
End
Attribute VB_Name = "frm_mainmenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub mnu_exit_Click()
    Unload Me
End Sub

Private Sub mnu_first_Click()
    load_first_sec_type = "F"
    frm_weighbridge_entry.Show
    frm_weighbridge_entry.ZOrder
End Sub

Private Sub mnu_import_Click()
    frm_import.Show
    frm_import.ZOrder
End Sub

Private Sub mnu_item_group_Click()
    
    frm_itemgroup_master.Show
    frm_itemgroup_master.ZOrder

''    frm_import.Show
''    frm_import.ZOrder

End Sub

Private Sub mnu_Material_Click()
    frm_item_master.Show
    frm_item_master.ZOrder
End Sub

Private Sub mnu_Party_Click()
    frm_party_master.Show
    frm_party_master.ZOrder
End Sub


Private Sub mnu_port_Click()
      PortSetting.Show
      PortSetting.ZOrder
End Sub

Private Sub mnu_second_Click()
      load_first_sec_type = "S"
      frm_weighbridge_entry.Show
      frm_weighbridge_entry.ZOrder
End Sub

Private Sub mnu_weightment_Details_Click()
    frm_reports.Show
    frm_reports.ZOrder
End Sub

Private Sub mu_duplicate_ticket_Click()
    frm_duplicate_ticket.Show
    frm_duplicate_ticket.ZOrder
End Sub
