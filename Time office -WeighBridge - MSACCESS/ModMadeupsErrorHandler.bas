Attribute VB_Name = "ModMadeupsErrorHandler"
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


