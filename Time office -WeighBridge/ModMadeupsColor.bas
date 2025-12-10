Attribute VB_Name = "ModMadeupsColor"
Option Explicit
Public gen_connection As ADODB.Connection
Public gen_connection_mysql As ADODB.Connection

Public gen_hrconnection As ADODB.Connection
Public gen_kgdl As ADODB.Connection
Public gen_kgdl_mdups As ADODB.Connection
Public gen_kgdl_denim As ADODB.Connection
Public gen_kgdl_itemcodeexport As ADODB.Connection
Public gen_mdups_color As New ADODB.Connection
Public form_type As Integer
Global chk As String
Public gst_finyear As String
Public gdt_date As Date
Public gboo_form_status As Boolean
Public gdt_fin_startdate As Date
Public gdt_fin_enddate As Date
Public gst_trantime As String
Public gin_finid As Integer
Public gst_Invtype As String
Public gst_LocalExportType As String * 1
Public gst_DirectIndirect_type As String
Public inv_type As String
Public export_no As Double
Public gboo_editoption As Boolean
Public gst_username As String
Public gin_usercode As Integer
Public gst_curcode As Integer
Public rg_flag As String
Global gst_password As String
Global CustType As String
Public Const gin_compcode = 4
Public gint_menu As Integer

Public load_first_sec_type As String

Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Sub frm_colors(frm As Form)
    Dim i As Integer
    On Error GoTo err_handler
    frm.Top = 0
    frm.Left = 0
    If form_type = 1 Then
        frm.BackColor = RGB(187, 196, 222)
    ElseIf form_type = 2 Then
        frm.BackColor = RGB(255, 228, 196)
    ElseIf form_type = 3 Then
        frm.BackColor = RGB(224, 255, 255)
    End If
    For i = 0 To frm.Controls.Count - 1
        If form_type = 1 Then
            If TypeOf frm.Controls(i) Is Label Then
                frm.Controls(i).ForeColor = RGB(0, 51, 0)
                frm.Controls(i).FontBold = True
                frm.Controls(i).BackStyle = 0
                frm.Controls(i).Font = "arial"
                frm.Controls(i).FontSize = 9
            End If
            If TypeOf frm.Controls(i) Is Frame Then
                frm.Controls(i).BackColor = RGB(187, 196, 222)
                frm.Controls(i).FontBold = True
                frm.Controls(i).ForeColor = RGB(205, 6, 70)
                frm.Controls(i).Font = "arial"
                frm.Controls(i).FontSize = 9
            End If
            If TypeOf frm.Controls(i) Is PictureBox Then
                frm.Controls(i).BackColor = RGB(250, 235, 215)
            End If
            If TypeOf frm.Controls(i) Is OptionButton Then
                frm.Controls(i).BackColor = RGB(187, 196, 222)
                frm.Controls(i).ForeColor = RGB(0, 51, 0)
                frm.Controls(i).FontBold = True
                frm.Controls(i).Font = "arial"
                frm.Controls(i).FontSize = 9
            End If

        ElseIf form_type = 2 Then
            If TypeOf frm.Controls(i) Is Label Then
                frm.Controls(i).ForeColor = RGB(139, 0, 0)
                frm.Controls(i).FontBold = True
                frm.Controls(i).BackStyle = 0
                frm.Controls(i).Font = "Arial"
                frm.Controls(i).FontSize = 9
            End If
            If TypeOf frm.Controls(i) Is Frame Then
                frm.Controls(i).BackColor = RGB(255, 228, 196)
                frm.Controls(i).FontBold = True
                frm.Controls(i).ForeColor = &HFF0000
                frm.Controls(i).Font = "Arial"
                frm.Controls(i).FontSize = 9
            End If
            If TypeOf frm.Controls(i) Is PictureBox Then
                frm.Controls(i).BackColor = RGB(143, 188, 139)
            End If
            If TypeOf frm.Controls(i) Is OptionButton Then
                frm.Controls(i).BackColor = RGB(255, 228, 196)
                frm.Controls(i).ForeColor = RGB(153, 50, 204)
                frm.Controls(i).FontBold = True
                frm.Controls(i).Font = "Arial"
                frm.Controls(i).FontSize = 9
            End If
            If TypeOf frm.Controls(i) Is CheckBox Then
                frm.Controls(i).BackColor = RGB(255, 228, 196)
                frm.Controls(i).ForeColor = RGB(153, 50, 204)
                frm.Controls(i).FontBold = True
                frm.Controls(i).Font = "Arial"
                frm.Controls(i).FontSize = 9
            End If
            If TypeOf frm.Controls(i) Is TextBox Then
                frm.Controls(i).BackColor = RGB(255, 250, 250)
                frm.Controls(i).Appearance = 0
                frm.Controls(i).FontSize = 9
                frm.Controls(i).Font = "Arial"
            End If
        ElseIf form_type = 3 Then
            If TypeOf frm.Controls(i) Is Label Then
                frm.Controls(i).ForeColor = RGB(0, 51, 0)
                frm.Controls(i).FontBold = True
                frm.Controls(i).BackStyle = 0
                frm.Controls(i).Font = "arial"
                frm.Controls(i).FontSize = 9
            End If
            If TypeOf frm.Controls(i) Is Frame Then
                frm.Controls(i).BackColor = RGB(224, 255, 255)
                frm.Controls(i).FontBold = True
                frm.Controls(i).ForeColor = RGB(95, 158, 160)
                frm.Controls(i).Font = "arial"
                frm.Controls(i).FontSize = 9
            End If
            If TypeOf frm.Controls(i) Is PictureBox Then
                frm.Controls(i).BackColor = RGB(216, 191, 216)
            End If
            If TypeOf frm.Controls(i) Is OptionButton Then
                frm.Controls(i).BackColor = RGB(224, 255, 255)
                frm.Controls(i).ForeColor = RGB(25, 25, 112)
                frm.Controls(i).FontBold = True
                frm.Controls(i).Font = "arial"
                frm.Controls(i).FontSize = 9
            End If
            If TypeOf frm.Controls(i) Is CheckBox Then
                frm.Controls(i).BackColor = RGB(224, 255, 255)
                frm.Controls(i).ForeColor = RGB(25, 25, 112)
                frm.Controls(i).FontBold = True
                frm.Controls(i).Font = "Arial"
                frm.Controls(i).FontSize = 9
            End If
        End If
        If TypeOf frm.Controls(i) Is TextBox Then
            frm.Controls(i).BackColor = RGB(255, 250, 250)
            frm.Controls(i).FontSize = 9
            frm.Controls(i).Font = "Arial"
            frm.Controls(i).Appearance = 1
            frm.Controls(i).BorderStyle = 1
        End If
        If TypeOf frm.Controls(i) Is ComboBox Then
            frm.Controls(i).BackColor = RGB(245, 255, 250)
            frm.Controls(i).Appearance = 0
            frm.Controls(i).FontSize = 9
            frm.Controls(i).Font = "Arial"
        End If
        If TypeOf frm.Controls(i) Is CommandButton Then
            frm.Controls(i).BackColor = &H8000000F
            frm.Controls(i).FontSize = 9
            frm.Controls(i).Font = "Arial"
            frm.Controls(i).FontBold = True
        End If
    Next i
    Exit Sub
err_handler:
    If gen_Validation(Err.Number, Err.Description) = 1 Then Resume
End Sub
Public Function cmb_maxlength(cmb_max As Integer, cmb As ComboBox, KeyValue As Integer)
    On Error GoTo err_handler
    If Len(cmb.Text) >= cmb_max Then
        cmb_maxlength = 0
        Exit Function
    End If
    cmb_maxlength = Asc(UCase(Chr(KeyValue)))
    Exit Function
err_handler:
    If gen_Validation(Err.Number, Err.Description) = 1 Then Resume
End Function
Public Sub gen_dbconnection()
    On Error GoTo err_handler
    Dim AdoCmd_getdate As New ADODB.Command
    Dim rs_getdate As New ADODB.Recordset
    
    Dim strcnn As String
    Dim strcnn_mysql As String

    Dim pst_conn As String, pst_ret As Integer
    pst_conn = Space(150)

    
     strcnn_mysql = "DRIVER={MySQL ODBC 3.51 Driver}; SERVER=10.0.0.251; PORT = 3306; DATABASE=shvpm; USER=root; PASSWORD=P@ssw0rD; OPTION=3; CHARSET = UTF8; SOCKET = MYSQL"
    'madeups dfd
    
    
    ''Set gen_connection = New ADODB.Connection
    ''gen_connection.CursorLocation = adUseClient
    ''gen_connection.Open strcnn
    
    Set gen_connection_mysql = Nothing
    Set gen_connection_mysql = New ADODB.Connection
    gen_connection_mysql.CursorLocation = adUseClient
    gen_connection_mysql.Open strcnn_mysql
    
   Set rs_getdate = Nothing
    Exit Sub
err_handler:
MsgBox ("Error")
Exit Sub

    If gen_Validation(Err.Number, Err.Description) = 1 Then Resume
End Sub
Public Sub cmb_popopcombine(cmb_name As Object, sql As String, Optional newin As Integer)
    Dim rs_popup  As New ADODB.Recordset
    Dim Flag As Boolean
    Flag = False
    On Error GoTo err_handler
    gen_connection.CommandTimeout = 1000
    rs_popup.Open sql, gen_connection, adOpenKeyset, adLockOptimistic
    cmb_name.Clear
    If rs_popup.RecordCount > 0 Then
        Flag = True
        cmb_name.Clear
        Do
            If newin = 1 Then
                cmb_name.AddItem Trim(rs_popup(0))
            Else
                cmb_name.AddItem Trim(rs_popup(1))
            End If
            If newin <> 1 Then
                cmb_name.ItemData(cmb_name.NewIndex) = rs_popup(0)
            End If
            rs_popup.MoveNext
        Loop Until rs_popup.EOF
    Else
        Flag = False
    
        End If
    Set rs_popup = Nothing
    Exit Sub
err_handler:
'Resume
    If gen_Validation(Err.Number, Err.Description) = 1 Then Resume
End Sub

Public Sub Fill_Combo(qry As String, obj_fill As Object, Optional objtype As Integer)
On Error GoTo err_handler:
    Dim rs_fillcombo As New ADODB.Recordset
    Dim pcmd_fillcombo As New ADODB.Command
    pcmd_fillcombo.ActiveConnection = gen_connection
    pcmd_fillcombo.CommandText = qry
    Set rs_fillcombo = pcmd_fillcombo.Execute
    With rs_fillcombo
       ' .Open Qry, gcn_servall, adOpenStatic, adLockReadOnly
        obj_fill.Clear
        If Not .EOF Then
            While Not .EOF
                obj_fill.AddItem .Fields(1)
                If .Fields.Count > 1 Then obj_fill.ItemData(obj_fill.NewIndex) = .Fields(0)
                .MoveNext
            Wend
                obj_fill.ListIndex = 0
        End If
        .Close
    End With
    Set rs_fillcombo = Nothing
    Exit Sub
err_handler:
'Resume
    If gen_Validation(Err.Number, Err.Description) = 1 Then Resume
End Sub
Public Sub Message(i As Integer, Optional name As String)
    Dim msg(100)
    On Error GoTo err_handler
    msg(1) = " Record Not Found"
    msg(2) = " Already Exist"
    msg(3) = " Should Not Be Empty"
    msg(4) = " Record Saved Successfully"
    msg(5) = " Record Not Saved"
    msg(6) = " Select From List Box"
    msg(7) = " Record Successfully Updated"
    msg(8) = " Select From Combo Box "
    msg(9) = "    "
    MsgBox name & msg(i), , "SALES"
    Exit Sub
err_handler:
    If gen_Validation(Err.Number, Err.Description) = 1 Then Resume
End Sub
Public Sub frm_clear(frm As Form)
    Dim i As Integer
    On Error GoTo err_handler
    For i = 0 To frm.Controls.Count - 1
        If TypeOf frm.Controls(i) Is TextBox Then
            frm.Controls(i).Text = ""
        End If
        If TypeOf frm.Controls(i) Is ComboBox Or TypeOf frm.Controls(i) Is ListBox Then
            frm.Controls(i).Clear
        End If
        If TypeOf frm.Controls(i) Is CheckBox Then
            frm.Controls(i).Value = 0
        End If

    Next i
    Exit Sub
err_handler:
    If gen_Validation(Err.Number, Err.Description) = 1 Then Resume
End Sub
Public Function Numeric_Chk(key As Integer, txt_value As Variant, tot_len As Integer, Optional i As Integer, Optional d As Integer) As Integer
    On Error GoTo err_handler
    If key = vbKeyBack Then
        Numeric_Chk = key
        Exit Function
    End If
    If Len(txt_value) > tot_len Then
        Numeric_Chk = 0
        Exit Function
    End If
    If i = 0 Then
        If key = 46 Then
            If InStr(1, txt_value, ".") > 0 Then
                Numeric_Chk = 0
                Exit Function
            End If
        End If
    Else
        If InStr(1, txt_value, ".") = 0 Then
            If Len(txt_value) = i Then
                Numeric_Chk = 46
                Exit Function
            ElseIf key = 46 And Len(txt_value) < i Then
                Numeric_Chk = key
                Exit Function
            End If
        End If
    End If
    If i > 0 And d > 0 Then
        If key <> 46 And Len(txt_value) > i And InStr(1, txt_value, ".") = 0 Then
            Numeric_Chk = 8
            Exit Function
        End If
    End If
    
    If d > 0 Then
        If InStr(1, txt_value, ".") <> 0 Then
            If Len(Mid(txt_value, InStr(1, txt_value, ".") + 1, Len(txt_value))) >= d Then
                 Numeric_Chk = 0
                 Exit Function
            End If
        End If
    End If
    If Not (key > 47 And key < 58) And key <> 8 Then
        Numeric_Chk = 0
        Exit Function
    End If
    Numeric_Chk = key
    Exit Function
err_handler:
    If gen_Validation(Err.Number, Err.Description) = 1 Then Resume
End Function
Public Function Cabs(key As Integer) As Integer
    On Error GoTo err_handler
    Cabs = Asc(UCase(Chr(key)))
    Exit Function
err_handler:
    If gen_Validation(Err.Number, Err.Description) = 1 Then Resume
End Function
Public Function Mail_Id(Text1 As String)
    Dim pin_atloc As String
    On Error GoTo err_handler
    If Len(Trim(Text1)) = 0 Then Exit Function
    If Right(Trim(Text1), 1) = "@" And Right(Trim(Text1), 1) = "." Then
        MsgBox "Invalid Mail Id", vbInformation, "Message"
        Mail_Id = True
        Exit Function
    End If
    If Mid(Trim(Text1), 1, 1) = "@" Then
        MsgBox "Mail Id should not starts with @ symbol", vbInformation, "Message"
        Mail_Id = True
        Exit Function
    End If
    If InStr(1, Trim(Text1), "@") <= 0 Then
        MsgBox "Mail Id should contain @ symbol", vbInformation, "Message"
        Mail_Id = True
        Exit Function
    End If
    If InStr(1, Trim(Text1), "@") = Len(Trim(Text1)) Then
        MsgBox "Invalid Mail Id", vbInformation, "Message"
        Mail_Id = True
        Exit Function
    End If
    If InStr(1, Trim(Text1), ".") = Len(Trim(Text1)) Then
        MsgBox "Invalid Mail Id", vbInformation, "Message"
        Mail_Id = True
        Exit Function
    End If
    
    pin_atloc = InStr(1, Trim(Text1), "@")
    If InStr(pin_atloc + 1, Trim(Text1), "@") > 0 Then
        MsgBox "Invalid Mail Id", vbInformation, "Message"
        Mail_Id = True
        Exit Function
    End If
    If InStr(pin_atloc + 1, Trim(Text1), ".") <= 0 Then
        MsgBox "Invalid Mail Id", vbInformation, "Message"
        Mail_Id = True
        Exit Function
    End If
    If pin_atloc + 1 = InStr(1, Trim(Text1), ".") Then
        MsgBox "Invalid Mail Id", vbInformation, "Message"
        Mail_Id = True
        Exit Function
    End If
    Exit Function
err_handler:
    If gen_Validation(Err.Number, Err.Description) = 1 Then Resume
End Function
Public Function find_index_item_data(pcmb_source As ComboBox, pst_search As String)
    Dim pin_idx As Integer, chk As Integer
    On Error GoTo err_handler
    For pin_idx = 0 To pcmb_source.ListCount - 1
       If pcmb_source.ItemData(pin_idx) = Val(pst_search) Then
            find_index_item_data = pin_idx
            Exit Function
        End If
    Next
       find_index_item_data = -1
    Exit Function
err_handler:
    chk = gen_Validation(Err.Number, Err.Description)
    If chk = 1 Then Resume
End Function
Public Function find_index_text(pcmb_source As ComboBox, pst_search As String)
    Dim pin_idx As Integer
    On Error GoTo err_handler
    For pin_idx = 0 To pcmb_source.ListCount - 1
        If pcmb_source.List(pin_idx) = pst_search Then
        find_index_text = pin_idx
        Exit Function
        End If
    Next
    find_index_text = -1
    Exit Function
err_handler:
    If gen_Validation(Err.Number, Err.Description) = 1 Then Resume
End Function
Public Function chk_date(rdate As Date, fyear As String) As Boolean
    Dim fdate As Date
    Dim tdate As Date
    On Error GoTo err_handler
    fdate = Format("01/04/" + Mid(fyear, 1, 4), "dd/mm/yyyy")
    tdate = Format("31/03/" + Mid(fyear, 6, 9), "dd/mm/yyyy")
    If Not (CDate(fdate) <= rdate And CDate(tdate) >= rdate) Then
        MsgBox "The Date Must be Entered between  " + Format(fdate, "dd/mm/yyyy") + " - " + Format(tdate, "dd/mm/yyyy"), vbOKOnly
         chk_date = False
    Else
       If CDate(rdate) > CDate(gdt_date) Then
            MsgBox "The Date Should Not Be Greater Than  Current Date", vbInformation, "INVALID DATE"
            chk_date = False
       Else
            chk_date = True
       End If
    End If
    Exit Function
err_handler:
    If gen_Validation(Err.Number, Err.Description) = 1 Then Resume
End Function

Public Function chk_from_to_date(date_from As Date, date_to As Date) As Boolean
    If date_from > date_to Then
        MsgBox "From Date Should not be Greater than  To Date", vbOKOnly + vbInformation, "Inspection Report"
        chk_from_to_date = True
    Else
        chk_from_to_date = False
    End If
End Function

Public Sub chk_keyascii(chk_val As TextBox, fld_type As String, fld_width As Integer, fld_decimal As Integer, ByRef chk_keyascii As Integer)
On Error GoTo err_handler:
    Dim m_IntPart As Integer
    Dim m_DecPart As Integer
    If chk_val.SelLength > 0 Then chk_keyascii = 0
    If chk_val.SelStart = 0 And chk_keyascii = 8 Then
        chk_keyascii = 0
        Exit Sub
    End If
    m_IntPart = fld_width
    If chk_keyascii = vbKeyBack Then
        If Len(Trim(chk_val)) <= 0 Then Exit Sub
        If InStr(1, Mid(chk_val, chk_val.SelStart, IIf(chk_val.SelLength > 0, chk_val.SelLength, 1)), ".") > 0 Then
            chk_val = Mid(chk_val, 1, chk_val.SelStart - 1)
            chk_keyascii = 0
            chk_val.SelStart = Len(chk_val)
        End If
        Exit Sub
    End If
    Select Case fld_type
    Case "N", "D"
        m_DecPart = fld_decimal
        If chk_keyascii = Asc(".") And m_DecPart = 0 Then chk_keyascii = 0
        If chk_keyascii = Asc(".") And InStr(1, chk_val, ".") = 0 Then Exit Sub
        If (chk_keyascii < Asc("0") Or chk_keyascii > Asc("9")) Then
            chk_keyascii = 0
            Exit Sub
        End If
        'Integer part Checking when no decimal point
        If InStr(1, chk_val.Text, ".") < 1 Then
            If Len(chk_val) < m_IntPart Then
                Exit Sub
            Else
                chk_keyascii = 0
            End If
        End If
        If InStr(1, chk_val.Text, ".") > 0 Then
            If chk_val.SelStart < InStr(1, chk_val, ".") Then
            'Integer part Checking when having decimal point
                If InStr(1, chk_val, ".") - 1 >= m_IntPart Then
                    chk_keyascii = 0
                End If
            Else
            'Decimal part Checking
                If m_DecPart < Len(chk_val) - InStr(1, chk_val, ".") + 1 Then
                    chk_keyascii = 0
                End If
            End If
        End If
    Case "V"
        m_IntPart = fld_width
        If chk_keyascii = 39 Then
            chk_keyascii = 0
        End If
        If chk_keyascii <> vbKeyBack And Len(chk_val.Text) >= m_IntPart Then
            chk_keyascii = 0
        End If
    End Select
    Exit Sub
err_handler:
    If gen_Validation(Err.Number, Err.Description) = 1 Then Resume
End Sub
Public Function find_subgroup(parm_grp As String)
On Error GoTo err_handler
    Dim adocmd_subgroup As ADODB.Command
    Dim rs_subgroup As ADODB.Recordset, pst_grpcode1, pst_grpcode, i As Integer
    pst_grpcode = parm_grp
    pst_grpcode1 = parm_grp
    Set adocmd_subgroup = New ADODB.Command
    Set rs_subgroup = New ADODB.Recordset
    adocmd_subgroup.ActiveConnection = gen_connection

    Do
        adocmd_subgroup.CommandText = " select grp_code,grp_name from acc_group_master where " & _
                                        " grp_parent_code in (" & parm_grp & " ) " & _
                                        " and grp_comp_code = " & gin_compcode
                                        
        adocmd_subgroup.CommandType = adCmdText
        Set rs_subgroup = adocmd_subgroup.Execute
        parm_grp = ""
If rs_subgroup.RecordCount > 0 Then
            For i = 1 To rs_subgroup.RecordCount
                pst_grpcode1 = Trim(pst_grpcode1) & "," & Trim(rs_subgroup("grp_code"))
                parm_grp = Trim(parm_grp) & "," & rs_subgroup("grp_code")
                rs_subgroup.MoveNext
            Next i
            parm_grp = Mid(parm_grp, 2)
            pst_grpcode1 = Trim(pst_grpcode1)
        End If
Loop Until Mid(parm_grp, 2, 1) = ""
    find_subgroup = pst_grpcode1
    Exit Function
err_handler:
    If gen_Validation(Err.Number, Err.Description) = 1 Then Resume
End Function


