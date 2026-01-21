VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frm_import 
   Caption         =   "Form1"
   ClientHeight    =   5865
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8295
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5865
   ScaleWidth      =   8295
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmd_import_suppliers 
      Caption         =   "IMPORT - suppliers"
      Height          =   735
      Left            =   4440
      TabIndex        =   4
      Top             =   3360
      Width           =   1815
   End
   Begin VB.Frame Frame4 
      Height          =   1095
      Left            =   1440
      TabIndex        =   1
      Top             =   1200
      Width           =   6255
      Begin MSComCtl2.DTPicker repDate 
         Height          =   375
         Left            =   3360
         TabIndex        =   2
         Top             =   360
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   116785153
         CurrentDate     =   44565
      End
      Begin VB.Label Label10 
         Caption         =   "DOWNLOAD FOR THE DATE"
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
         Left            =   240
         TabIndex        =   3
         Top             =   480
         Width           =   3015
      End
   End
   Begin VB.CommandButton cmd_import 
      Caption         =   "IMPORT"
      Enabled         =   0   'False
      Height          =   735
      Left            =   1920
      TabIndex        =   0
      Top             =   3360
      Width           =   1815
   End
End
Attribute VB_Name = "frm_import"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mwt As Integer
Dim adocmd_mysql As New ADODB.Command
Dim adors As New ADODB.Recordset
Dim pst_qry, mdb_qry As String

 'Private Declare Sub GenerateBMP _
''                Lib "C:\WINDOWS\system32\quricol32.dll" _
''                Alias "GenerateBMPW" ( _
''                    ByVal FileName As Long, _
''                ByVal Text As Long, _
''                ByVal Margin As Long, _
''                ByVal Size As Long, _
''                ByVal Level As TErrorCorretion)
Private Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)
    
Dim compcode, fincode As Integer

Private Sub cmd_import_Click()

     '' strcnn_mysql = "DRIVER={MySQL ODBC 3.51 Driver}; SERVER=10.0.0.251; PORT = 3306; DATABASE=shvpm; USER=root; PASSWORD=P@ssw0rD; OPTION=3; CHARSET = UTF8; SOCKET = MYSQL"
    
    'madeups dfd
    
    
    ''Set gen_connection = New ADODB.Connection
    ''gen_connection.CursorLocation = adUseClient
    ''gen_connection.Open strcnn
    
''    Set gen_connection_mysql = Nothing
''    Set gen_connection_mysql = New ADODB.Connection
''    gen_connection_mysql.CursorLocation = adUseClient
''    gen_connection_mysql.Open strcnn_mysql
        
        
    Dim adocmd_mysql As New ADODB.Command
    Dim adors As New ADODB.Recordset
    adocmd_mysql.ActiveConnection = gen_connection_mysql

    Dim dsnmdb As String
    
    
    Dim mdbrs As New ADODB.Recordset
    
       dsnmdb = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=\\10.0.0.252\vbexe\trucks.mdb"
     
       dsnmdb = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=d:\trucks.mdb"
       
''     mdb_qry = "Select *  from Tickets where Date between #" & Format(repDate, "MM/dd/yyyy") & "# and  #" & Format(end_date + 1, "MM/dd/yyyy") & "# order by TicketNumber"
''    mdb_qry = "Select *  from Tickets where Date = #" & Format(repDate.Value, "MM/dd/yyyy") & "# and  TicketNumber >= " & Val(txtTicketNo.Text) & " and State = 'Online Second Transaction'  order by TicketNumber"
    mdb_qry = "Select *  from Tickets where Date = #" & Format(repDate.Value, "MM/dd/yyyy") & "# and State = 'Online Second Transaction' order by TicketNumber"
    mdbrs.Open mdb_qry, dsnmdb, adOpenDynamic, adLockOptimistic

    While Not mdbrs.EOF
 ''       MsgBox (CStr(mdbrs!TicketNumber) + " - " + mdbrs!VehicleNumber)
        
        
        pst_qry = "select * from trn_weight_card where wc_compcode = " & compcode & "  and wc_fincode = " & fincode & "  and wc_ticketno = " & mdbrs!TicketNumber
        adocmd_mysql.CommandText = pst_qry
        Set adors = adocmd_mysql.Execute
        If adors.RecordCount = 0 Then

            pst_qry = "insert into trn_weight_card (wc_compcode, wc_fincode, wc_ticketno, wc_date,wc_time, wc_area_code, wc_sup_code, wc_item, wc_vehicleno, wc_emptywt, wc_loadwt, wc_netwt, wc_supplier,wc_acceptedwt) VALUES ( " & compcode & ", " & fincode & "," & mdbrs!TicketNumber & " ,'" & Format(mdbrs!LoadWeightDate, "yyyy-MM-dd") & "','" & mdbrs!LoadWeightTime & "', 0,0,'" & mdbrs!MaterialName & "' ,'" & mdbrs!VehicleNumber & "' ," & mdbrs!EmptyWeight & "," & mdbrs!LoadedWeight & "," & mdbrs!NetWeight & ",'" & mdbrs!SupplierName & "', " & mdbrs!NetWeight & ")"
            adocmd_mysql.CommandTimeout = 300
            adocmd_mysql.CommandText = pst_qry
            adocmd_mysql.Execute pst_qry
        End If

        mdbrs.MoveNext
    Wend
    mdbrs.Close
    
    MsgBox ("Data Imported Completed..")
   
End Sub

Private Sub Command1_Click()

End Sub

Private Sub cmd_import_suppliers_Click()
    Dim adocmd_mysql As New ADODB.Command
    Dim adors As New ADODB.Recordset
    adocmd_mysql.ActiveConnection = gen_connection_mysql

    Dim dsnmdb As String
    
    
    Dim mdbrs As New ADODB.Recordset
    
       dsnmdb = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=\\10.0.0.252\vbexe\trucks.mdb"
     
       dsnmdb = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=d:\trucks.mdb"
       
''     mdb_qry = "Select *  from Tickets where Date between #" & Format(repDate, "MM/dd/yyyy") & "# and  #" & Format(end_date + 1, "MM/dd/yyyy") & "# order by TicketNumber"
''    mdb_qry = "Select *  from Tickets where Date = #" & Format(repDate.Value, "MM/dd/yyyy") & "# and  TicketNumber >= " & Val(txtTicketNo.Text) & " and State = 'Online Second Transaction'  order by TicketNumber"
    mdb_qry = "Select *  from Suppliers"
    mdbrs.Open mdb_qry, dsnmdb, adOpenDynamic, adLockOptimistic

    While Not mdbrs.EOF
        

            pst_qry = "insert into mas_wb_party (party_code, party_name) VALUES ( '" & mdbrs!SupplierCode & "' , '" & mdbrs!SupplierName & "')"
            adocmd_mysql.CommandTimeout = 300
            adocmd_mysql.CommandText = pst_qry
            adocmd_mysql.Execute pst_qry

        mdbrs.MoveNext
    Wend
    mdbrs.Close
    
    MsgBox ("Data Imported Completed..")

End Sub

Private Sub Form_Load()
    Call gen_dbconnection
      
    repDate.Value = Now
    compcode = 1
    fincode = 23
    
End Sub

