Attribute VB_Name = "payroll"
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Any) As Long
Global paydb As New ADODB.Connection
Global payrs As New ADODB.Recordset
Global paydb2 As New ADODB.Connection
Global payrs2 As New ADODB.Recordset
Global payquery As String
Global pay As String
Global sql As String
Global sql2 As String
Global sql3 As String
Global sql4 As String
Global sql5 As String
Global millname As String
Global name1 As String
Global name2 As String
Global name3 As String
Global name4 As String
Global name5 As String
Global menuchk As Byte
Global savechk As Byte
Global oldtext As String
Global code As Integer
Global dname As String
Global dcode As Integer
Global dept_name As String
Global employee_name As String
Global attendance_staus As String
Global emptypecode As Integer
Global mdate As Date
Global endrow As Byte
Global pay_calchk As Integer
Global emptype_chk As Integer
Global company_code As String
Global at_rep_opt As Integer
Global hod As Boolean
Global att_dat As Date
Global attstatus As String
Global finyear As Integer
Private Sub main()
''    pay = "Provider=SQLOLEDB.1;Password=serdat;Persist Security Info=True;User ID=sa;DATABASE=spl_others;Data Source=spplserver"
    pay = "Provider=SQLOLEDB.1;Password=serdat;Persist Security Info=True;User ID=sa;DATABASE=anna_test;Data Source=servalldata"
        
End Sub
Public Function gen_Validation(errno As Long, errdes As String) As Integer
Dim Pst_Msg As String
gen_Validation = 0
Select Case errno
       Case -2147467259
            Pst_Msg = "Invalid DSN/MDB/DB Name"
       Case -2147217900
            Pst_Msg = "Query contains single quotes"
       Case -2147217884
            Pst_Msg = "Invalid Operation"
       Case -2147217873
            Pst_Msg = "Violation of Key Constraints"
       Case 5
            Pst_Msg = "Invalid Procedure"
       Case 6
            Pst_Msg = "Value Overflow"
       Case 7
            Pst_Msg = "Memory Low"
       Case 9
            Pst_Msg = "Invalid Array Index"
       Case 11
            Pst_Msg = "Some of the values contains zero"
       Case 13
            Pst_Msg = "Invalid Type"
       Case 18
            Pst_Msg = "Interrupt by User"
       Case 20
            Pst_Msg = "Unable to Resume"
       Case 28
            Pst_Msg = "Internal Error"
       Case 48
            Pst_Msg = "Specified DLL not found"
       Case 49
            Pst_Msg = "Improper DLL call"
       Case 51
            Pst_Msg = "Internal Error"
       Case 52
            Pst_Msg = "Invalid File Name or Number"
       Case 53
            Pst_Msg = "File Not Found"
       Case 54
            Pst_Msg = "Invalid File Mode"
       Case 55
            Pst_Msg = "Operation not allowed"
       Case 57
            Pst_Msg = "Hardware Error"
       Case 58
            Pst_Msg = "File already exists"
       Case 63
            Pst_Msg = "Invalid Record Number"
       Case 67
            Pst_Msg = "Buffer overflow"
       Case 68
            Pst_Msg = "Hardware not Found"
       Case 70
            Pst_Msg = "Access Denied"
       Case 75
            Pst_Msg = "File Access Error"
       Case 76
            Pst_Msg = "Invalid Path"
       Case 91
            Pst_Msg = "Invalid Object"
            MsgBox Pst_Msg
            gen_Validation = 1
       Case 94
            Pst_Msg = "Inputs should not be Empty"
       Case 298
            Pst_Msg = "System DLL could not be loaded"
       Case 321
            Pst_Msg = "Invalid file format"
       Case 326
            Pst_Msg = "Invalid Resource Identifier"
       Case 337
            Pst_Msg = "Component does not exist"
       Case 340
            Pst_Msg = "Invalid array element"
       Case 360
            Pst_Msg = "Object already loaded"
       Case 364
            Pst_Msg = "Object already unloaded"
       Case 365
            Pst_Msg = "Context unable to Unload"
       Case 380
            Pst_Msg = "Invalid Property value"
       Case 381
            Pst_Msg = "Invalid property array index"
       Case 383
            Pst_Msg = "Read Only Property"
       Case 387
            Pst_Msg = "Unable to set Property"
       Case 394
            Pst_Msg = "Write Only Property"
       Case 400
            Pst_Msg = "This screen already Displayed"
       Case 419
            Pst_Msg = "Access Denied"
       Case 424
            Pst_Msg = "Object required"
       Case 425
            Pst_Msg = "Improper use of Object"
       Case 429
            Pst_Msg = "Component Error"
       Case 438
            Pst_Msg = "Property not Supported by this Object"
       Case 440
            Pst_Msg = "OLE Automation error"
       Case 445
            Pst_Msg = "Object does not support this action"
       Case 449
            Pst_Msg = "Argument should be given"
       Case 450
            Pst_Msg = "No.of Arguments mismatch"
       Case 461
            Pst_Msg = "Method or data member not found"
       Case 482
            Pst_Msg = "Printer error"
       Case 483
            Pst_Msg = "Printer error"
       Case 484
            Pst_Msg = "Printer error"
       Case 1205
            Pst_Msg = "Dead lock"
       Case 3001
            Pst_Msg = "Invalid Parameters"
       Case 3021
            Pst_Msg = "End of Table reached"
       Case 3219
            Pst_Msg = "Operation is not allowed"
       Case 3251
            Pst_Msg = "Unable to perform this Operation"
       Case 3265
            Pst_Msg = "Item cannot be found"
       Case 3420
            Pst_Msg = "Invalid Object"
       Case 3421
            Pst_Msg = "Invalid Type value"
       Case 3704
            Pst_Msg = "Operation disabled"
       Case 3705
            Pst_Msg = "Object is already open"
       Case 3706
            Pst_Msg = "Invalid Provider"
       Case 3709
            Pst_Msg = "Object already Closed"
       Case 8000
            Pst_Msg = "Port already open"
       Case 8001
            Pst_Msg = "Timeout value is too little"
       Case 8002
            Pst_Msg = "Invalid Port Number"
       Case 8003
            Pst_Msg = "Invalid use of Property"
       Case 8004
            Pst_Msg = "Invalid use of Property"
       Case 8005
            Pst_Msg = "Port already open"
       Case 8010
            Pst_Msg = "Hardware Error"
       Case 8012
            Pst_Msg = "Device already Closed"
       Case 8013
            Pst_Msg = "Device already open"
       Case 8018
            Pst_Msg = "Invalid Operation"
       Case 10004
            Pst_Msg = "Connection Error"
       Case 10018
            Pst_Msg = "Network Error"
       Case 10024
            Pst_Msg = "Timeout Error"
       Case 10026
            Pst_Msg = "Invalid Network Connection Mode"
       Case 10058
            Pst_Msg = "Bulk Copy not allowed"
       Case 10103
            Pst_Msg = "I/O Error"
       Case 30001
            Pst_Msg = "Cannot create button"
       Case 30002
            Pst_Msg = "Cannot create a timer resource"
       Case 32002
            Pst_Msg = "File does not exists"
       Case 32005
            Pst_Msg = "Invalid or missing key name."
       Case 32010
            Pst_Msg = "Invalid object."
       Case 32012
            Pst_Msg = "Unable to create object."
       Case 35773
            Pst_Msg = "Invalid Date"
       Case Else
            Pst_Msg = errdes
End Select
If InStr(1, LCase(errdes), "timeout expired") > 0 Then
    gen_Validation = 1
Else
    MsgBox Pst_Msg, vbInformation, "Error"
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
    chk = gen_Validation(Err.Number, Err.Description)
    If chk = 1 Then Resume
End Sub
Function find_deptname(data_code As String) As String
    dname = ""
    sql = ("Select * from pdept_mas where dept_code = '" & data_code & "'")
    Set paydb2 = New ADODB.Connection
    Set payrs2 = New ADODB.Recordset
    paydb2.Open pay
    payrs2.Open sql, paydb2, adOpenDynamic, adLockOptimistic
    If Not payrs2.EOF Then
       dname = payrs2(1)
    End If
    Exit Function
End Function

Function find_empdetails(data_code As String) As String
    sql = ("Select * from emp_mas where emp_idcode = " & data_code & " and emp_company = '" & company_code & "'")
    Set paydb2 = New ADODB.Connection
    Set payrs2 = New ADODB.Recordset
    paydb2.Open pay
    payrs2.Open sql, paydb2, adOpenDynamic, adLockOptimistic
    If Not payrs2.EOF Then
       employee_name = payrs2("emp_name")
       emptypecode = payrs2.Fields("emp_type")
       find_deptname (payrs2.Fields("emp_dept"))
       dept_name = dname
    End If
    Exit Function
End Function

Function find_desiname(data_code As String) As String
    sql = ("Select * from pdesi_mas where pdesi_code = '" & data_code & "'")
    Set paydb2 = New ADODB.Connection
    Set payrs2 = New ADODB.Recordset
    paydb2.Open pay
    payrs2.Open sql, paydb2, adOpenDynamic, adLockOptimistic
    If Not payrs2.EOF Then
       dname = payrs2(1)
    End If
    Exit Function
End Function

Function find_etypename(data_code As String) As String
    sql = ("Select * from pemptype_mas where dtype_code = '" & data_code & "'")
    Set paydb2 = New ADODB.Connection
    Set payrs2 = New ADODB.Recordset
    paydb2.Open pay
    payrs2.Open sql, paydb2, adOpenDynamic, adLockOptimistic
    If Not payrs2.EOF Then
       dname = payrs2(1)
    End If
    Exit Function
End Function

Function find_qualifyname(data_code As String) As String
    sql = ("Select * from pqly_mas where pqly_code = '" & data_code & "'")
    Set paydb2 = New ADODB.Connection
    Set payrs2 = New ADODB.Recordset
    paydb2.Open pay
    payrs2.Open sql, paydb2, adOpenDynamic, adLockOptimistic
    If Not payrs2.EOF Then
       dname = payrs2(1)
    End If
    Exit Function
End Function

Function find_religioncode(data_code As String) As String
    sql = ("Select * from preli_mas where preli_name = '" & data_code & "'")
    Set paydb2 = New ADODB.Connection
    Set payrs2 = New ADODB.Recordset
    paydb2.Open pay
    payrs2.Open sql, paydb2, adOpenDynamic, adLockOptimistic
    If Not payrs2.EOF Then
       dcode = payrs2(0)
    End If
    Exit Function
End Function
Function find_communitycode(data_code As String) As String
    sql = ("Select * from pcomm_mas where pcomm_name = '" & data_code & "'")
    Set paydb2 = New ADODB.Connection
    Set payrs2 = New ADODB.Recordset
    paydb2.Open pay
    payrs2.Open sql, paydb2, adOpenDynamic, adLockOptimistic
    If Not payrs2.EOF Then
       dcode = payrs2(0)
    End If
    Exit Function
End Function

Function find_castecode(data_code As String) As String
    sql = ("Select * from pcast_mas where pcast_name = '" & data_code & "'")
    Set paydb2 = New ADODB.Connection
    Set payrs2 = New ADODB.Recordset
    paydb2.Open pay
    payrs2.Open sql, paydb2, adOpenDynamic, adLockOptimistic
    If Not payrs2.EOF Then
       dcode = payrs2(0)
    End If
    Exit Function
End Function
Function find_deptcode(data_code As String) As String
    sql = ("Select * from pdept_mas where dept_name = '" & data_code & "'")
    Set paydb2 = New ADODB.Connection
    Set payrs2 = New ADODB.Recordset
    paydb2.Open pay
    payrs2.Open sql, paydb2, adOpenDynamic, adLockOptimistic
    If Not payrs2.EOF Then
       dcode = payrs2(0)
    End If
    Exit Function
End Function
Function find_typecode(data_code As String) As String
    sql = ("Select * from pemptype_mas where dtype_name = '" & data_code & "'")
    Set paydb2 = New ADODB.Connection
    Set payrs2 = New ADODB.Recordset
    paydb2.Open pay
    payrs2.Open sql, paydb2, adOpenDynamic, adLockOptimistic
    If Not payrs2.EOF Then
       dcode = payrs2(0)
    End If
    Exit Function
End Function
Function find_designcode(data_code As String) As String
    sql = ("Select * from pdesi_mas where pdesi_name = '" & data_code & "'")
    Set paydb2 = New ADODB.Connection
    Set payrs2 = New ADODB.Recordset
    paydb2.Open pay
    payrs2.Open sql, paydb2, adOpenDynamic, adLockOptimistic
    If Not payrs2.EOF Then
       dcode = payrs2(0)
    End If
    Exit Function
End Function
Function find_qualifycode(data_code As String) As String
    sql = ("Select * from pqly_mas where pqly_name = '" & data_code & "'")
    Set paydb2 = New ADODB.Connection
    Set payrs2 = New ADODB.Recordset
    paydb2.Open pay
    payrs2.Open sql, paydb2, adOpenDynamic, adLockOptimistic
    If Not payrs2.EOF Then
       dcode = payrs2(0)
    End If
    Exit Function
End Function
Function find_attnstatus(data_code As String) As String
    sql = ("Select * from attn_status_mas where attn_type_code = '" & data_code & "'")
    Set paydb2 = New ADODB.Connection
    Set payrs2 = New ADODB.Recordset
    paydb2.Open pay
    payrs2.Open sql, paydb2, adOpenDynamic, adLockOptimistic
    If Not payrs2.EOF Then
       attendance_staus = payrs2(1)
    End If
    Exit Function
End Function
Function find_attn_status_code(data_code As String) As String
    sql = ("Select * from attn_status_mas where attn_type_name = '" & data_code & "'")
    Set paydb2 = New ADODB.Connection
    Set payrs2 = New ADODB.Recordset
    paydb2.Open pay
    payrs2.Open sql, paydb2, adOpenDynamic, adLockOptimistic
    If Not payrs2.EOF Then
       attendance_staus = payrs2(0)
    Else
       attendance_staus = 5
    End If
    Exit Function
End Function

Public Function Numeric_Chk(key As Integer, txt_value As Variant, tot_len As Integer, Optional i As Integer, Optional d As Integer) As Integer
    If key = vbKeyBack Then
        Numeric_Chk = key
        Exit Function
    End If
    If Len(txt_value) > tot_len + 1 Then
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
 '   If Len(txt_pckconv.Text) = 5 And Not Len(txt_pckconv.Text) < 4 And (KeyAscii > 47 And KeyAscii < 58) Then
 '   If Len(txt_pckconv.Text) = 5 And KeyAscii <> 46 And (KeyAscii > 47 And KeyAscii < 58) Then
 '        KeyAscii = 0
 '   End If
    If Not (key > 47 And key < 58) And key <> 8 And key <> 13 Then
        Numeric_Chk = 0
        Exit Function
    End If
    Numeric_Chk = key
    Exit Function
End Function
       
Public Function hours_Chk(key As Integer, txt_value As Variant, tot_len As Integer, Optional i As Integer, Optional d As Integer) As Integer
    If key = vbKeyBack Then
        hours_Chk = key
        Exit Function
    End If
    If Len(txt_value) = 0 And key = Asc(".") Then txt_value = "00.00"
    
    If Len(txt_value) > tot_len + 1 Then
        hours_Chk = 0
        Exit Function
    End If
    If i = 0 Then
        If key = 46 Then
            If InStr(1, txt_value, ".") > 0 Then
                hours_Chk = 0
                Exit Function
            End If
        End If
    Else
        If key > Asc("2") And Len(txt_value) = 0 Then txt_value = "00.00"
        If key > Asc("4") And Len(txt_value) = 1 And Left(txt_value, 1) = "2" Then txt_value = "20.00"
        If InStr(1, txt_value, ".") = 0 Then
            If Len(txt_value) = i Then
                hours_Chk = 46
                Exit Function
            ElseIf key = 46 And Len(txt_value) < i Then
                hours_Chk = key
                Exit Function
            End If
        End If
    End If
    ''    If i > 0 And d > 0 Then
''        If key <> 46 And Len(txt_value) > i And InStr(1, txt_value, ".") = 0 Then
''            hours_Chk = 8
''            Exit Function
''        End If
''    End If
    If Len(txt_value) = 2 And Mid(txt_value, 3, 1) <> "." Then Exit Function
        If d > 0 Then
        If Val(Left(txt_value, 2)) = 24 And key <> Asc("0") Then Exit Function
        If key > Asc("5") And Len(txt_value) = 3 Then Exit Function
        If InStr(1, txt_value, ".") <> 0 Then
            If Len(Mid(txt_value, InStr(1, txt_value, ".") + 1, Len(txt_value))) >= d Then
                 hours_Chk = 0
                 Exit Function
            End If
        End If
    End If
    If Not (key > 47 And key < 58) And key <> 8 And key <> 13 Then
        hours_Chk = 0
        Exit Function
    End If
    hours_Chk = key
    Exit Function
End Function
      
Public Function attndays_Chk(key As Integer, txt_value As Variant, tot_len As Integer, Optional i As Integer, Optional d As Integer) As Integer
    If key = vbKeyBack Then
        attndays_Chk = key
        Exit Function
    End If
    If Len(txt_value) = 0 And key = Asc(".") Then txt_value = "00.00"
    
    If Len(txt_value) > tot_len + 1 Then
        attndays_Chk = 0
        Exit Function
    End If
    If i = 0 Then
        If key = 46 Then
            If InStr(1, txt_value, ".") > 0 Then
                attndays_Chk = 0
                Exit Function
            End If
        End If
    Else
        If key > Asc("2") And Len(txt_value) = 0 Then txt_value = "00.00"
        If key > Asc("6") And Len(txt_value) = 1 And Left(txt_value, 1) = "2" Then txt_value = "20.00"
        If InStr(1, txt_value, ".") = 0 Then
            If Len(txt_value) = i Then
                attndays_Chk = 46
                Exit Function
            ElseIf key = 46 And Len(txt_value) < i Then
                attndays_Chk = key
                Exit Function
            End If
        End If
    End If
    ''    If i > 0 And d > 0 Then
''        If key <> 46 And Len(txt_value) > i And InStr(1, txt_value, ".") = 0 Then
''            hours_Chk = 8
''            Exit Function
''        End If
''    End If
    If Len(txt_value) = 2 And Mid(txt_value, 3, 1) <> "." Then Exit Function
        If d > 0 Then
        If Val(Left(txt_value, 2)) = 24 And key <> Asc("0") Then Exit Function
''        If key > Asc("5") And Len(txt_value) = 3 Then Exit Function
        If InStr(1, txt_value, ".") <> 0 Then
            If Len(Mid(txt_value, InStr(1, txt_value, ".") + 1, Len(txt_value))) >= d Then
                 attndays_Chk = 0
                 Exit Function
            End If
        End If
    End If
    If Not (key > 47 And key < 58) And key <> 8 And key <> 13 Then
        attndays_Chk = 0
        Exit Function
    End If
    attndays_Chk = key
    Exit Function
End Function
      
Public Function date_Chk(key As Integer, txt_value As Variant) As Integer
    If key = vbKeyBack Then
        date_Chk = key
        Exit Function
    End If
    If Len(txt_value) = 10 Then Exit Function
    If Len(Trim(txt_value)) = 3 Then
         If LTrim(Mid(txt_value, 3, 1)) <> "/" Then Exit Function
    End If
    If Len(Trim(txt_value)) = 6 Then
         If LTrim(Mid(txt_value, 6, 1)) <> "/" Then Exit Function
    End If
   
    If key < Asc("0") Or key > Asc("9") Then
        If key = Asc("/") Then
           date_Chk = key
        Else
           Exit Function
        End If
    Else
       date_Chk = key
    End If
End Function

Function find_present_status(ecode As Long) As String
    sql = ("Select * from tmp_attn where empcode = " & ecode & " and attndate ='" & Format(att_dat, "mm/dd/yyyy") & "'")
    Set paydb2 = New ADODB.Connection
    Set payrs2 = New ADODB.Recordset
    paydb2.Open pay
    payrs2.Open sql, paydb2, adOpenDynamic, adLockOptimistic
    If Not payrs2.EOF Then
       attstatus = "PRESENT"
    End If
    Exit Function
End Function
