Attribute VB_Name = "attn_module"
Global attndb As New ADODB.Connection
Global attnrs As New ADODB.Recordset
Global mdbrs As New ADODB.Recordset
Global constr As String
Global millcode, fincode As Integer
Global millname As String
Global sqlqry, mdbqry1, mdbqry2, dsnmdb As String

Public Sub Fill_Combo_mill(qry As String, obj_fill As Object, Optional objtype As Integer)
    Dim rs_fillcombo As New ADODB.Recordset
    With rs_fillcombo
        .Open qry, attndb, adOpenStatic, adLockReadOnly
        obj_fill.Clear
        If Not .EOF Then
            While Not .EOF
                obj_fill.AddItem .Fields(0)
                If .Fields.Count > 1 Then obj_fill.ItemData(obj_fill.NewIndex) = .Fields(1)
                .MoveNext
            Wend
            obj_fill.ListIndex = 0
        End If
        .Close
    End With
    Set rs_fillcombo = Nothing
End Sub

Public Sub Fill_Combo(qry As String, obj_fill As Object, Optional objtype As Integer)
    Dim rs_fillcombo As New ADODB.Recordset
    Dim reccount As Integer
    reccount = 0
    With rs_fillcombo
        .Open qry, attndb, adOpenStatic, adLockReadOnly
        obj_fill.Clear
        If Not .EOF Then
            While Not .EOF
                obj_fill.AddItem .Fields(0)
                If .Fields.Count > 1 Then obj_fill.ItemData(obj_fill.NewIndex) = .Fields(1)
                reccount = reccount + 1
                .MoveNext
            Wend
            obj_fill.ListIndex = reccount - 1
        End If
        .Close
    End With
    Set rs_fillcombo = Nothing
End Sub

